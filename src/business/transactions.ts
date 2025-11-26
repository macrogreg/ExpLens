/// <reference types="office-js" />

import { errorTypeMessageString, formatDateUtc } from "src/util/format_util";
import { ensureSheetActive, findTableByNameOnSheet, getRangeBasedOn } from "./excel-util";
import type { Transaction, TransactionColumnSpec, TransactionRowData, TransactionRowValue } from "./transaction-tools";
import {
    getTransactionColumnValue,
    LunchIdColumnName,
    createTransactionColumnsSpecs,
    tryGetTagGroupFromColumnName,
} from "./transaction-tools";
import { isNullOrWhitespace } from "src/util/string_util";
import { parseTag } from "./tags";
import { authorizedFetch } from "./fetch-tools";
import type * as Lunch from "./lunchmoney-types";
import { IndexedMap } from "./IndexedMap";
import type { SyncContext } from "./sync-driver";

export const SheetNameTransactions = "LM.Transactions";
const TableNameTransactions = "LM.TransactionsTable";

function isColumnNamingEquivalent(
    columnsSpecs: IndexedMap<string, TransactionColumnSpec>,
    actualColumnNames: string[]
) {
    let expC = 0,
        actC = 0;
    while (true) {
        // Skip tag columns during structural comparison:
        while (expC < columnsSpecs.length && tryGetTagGroupFromColumnName(columnsSpecs.getByIndex(expC)!.name)) {
            expC++;
        }
        while (actC < actualColumnNames.length && tryGetTagGroupFromColumnName(actualColumnNames[actC]!)) {
            actC++;
        }

        // If both comparison column lists are exhausted, they match:
        if (expC === columnsSpecs.length && actC === actualColumnNames.length) {
            return true;
        }

        // If only of of the comparison column lists is exhausted, they do not match:
        if (expC === columnsSpecs.length || actC === actualColumnNames.length) {
            console.error(
                `isColumnNamingEquivalent(..):\n` +
                    `The lengths of 'columnsSpecs' (${columnsSpecs.length}) and` +
                    ` actualColumnNames (${actualColumnNames.length}) are different after accounting` +
                    ` for dynamic tag columns.`
            );
            return false;
        }

        // If col names at current cursors are different, lists do not match:
        if (columnsSpecs.getByIndex(expC)?.name !== actualColumnNames[actC]) {
            console.error(
                `isColumnNamingEquivalent(..):\n` +
                    `After accounting for dynamic tag columns, aligned headers are not the same:\n` +
                    `columnsSpecs.getByIndex(${expC})?.name !== actualColumnNames[${actC}]` +
                    ` ('${columnsSpecs.getByIndex(expC)?.name}' !=== '${actualColumnNames[actC]}').`
            );
            return false;
        }

        // Names at current positions match, move cursors forward:
        expC++;
        actC++;
    }
}

async function createNewTranTable(
    tranColumnsSpecs: IndexedMap<string, TransactionColumnSpec>,
    sheet: Excel.Worksheet,
    context: Excel.RequestContext
): Promise<Excel.Table> {
    // Table location:
    const tranTableOffs = { row: 7, col: 1 };

    // No data is loaded yet. Only use the no-tag-group column for tags initially:
    const tranSpecColNames = tranColumnsSpecs.map((col) => col.name);

    // Clear the are where we are about to create the table:
    const tableInitRange = getRangeBasedOn(sheet, tranTableOffs, 0, 0, 2, tranSpecColNames.length);

    tableInitRange.clear();
    tableInitRange.conditionalFormats.clearAll();
    await context.sync();

    // Print column headers:
    getRangeBasedOn(sheet, tranTableOffs, 0, 0, 1, tranSpecColNames.length).values = [tranSpecColNames];

    // Create table:
    const table = sheet.tables.add(tableInitRange, true);
    table.name = TableNameTransactions;
    table.style = "TableStyleMedium9"; // e.g."TableStyleMedium2", "TableStyleDark1", "TableStyleLight9" ...

    await context.sync();

    // Load frequently used properties:
    table.load(["name", "id"]);
    table.getRange().load(["address"]);
    await context.sync();

    console.debug(`New Transactions table '${table.name}' created.`);
    return table;
}

export async function downloadTransactions(startDate: Date, endDate: Date, context: SyncContext) {
    // Find and activate the Transactions sheet:
    const tranSheet = await ensureSheetActive(SheetNameTransactions, context.excel);

    const errorMsgCell = tranSheet.getRange("B4");
    errorMsgCell.clear();
    await context.excel.sync();

    try {
        const tranColumnsSpecs: IndexedMap<string, TransactionColumnSpec> = createTransactionColumnsSpecs(context);

        // Is there an existing Transactions table?
        const prevTranTableInfo = await findTableByNameOnSheet(
            TableNameTransactions,
            context.sheets.trans,
            context.excel
        );

        // If there is no existing table, create an empty one:
        const tranTable =
            prevTranTableInfo === null
                ? await createNewTranTable(tranColumnsSpecs, tranSheet, context.excel)
                : prevTranTableInfo.table;

        // Get the expected column names:
        if (!tranColumnsSpecs.has(LunchIdColumnName)) {
            const preexistenceQualifier = prevTranTableInfo === null ? "newly created" : "pre-existing";
            throw new Error(
                `The table '${tranTable.name}' (${preexistenceQualifier}) does not contain` +
                    ` the expected column '${LunchIdColumnName}'.`
            );
        }

        // Load the column names actually present in the table:
        tranTable.columns.load("items");
        await context.excel.sync();
        for (const col of tranTable.columns.items) {
            col.load("name");
        }
        await context.excel.sync();

        // Validate that the actual column names match the spec:
        if (
            !isColumnNamingEquivalent(
                tranColumnsSpecs,
                tranTable.columns.items.map((col) => col.name.trim())
            )
        ) {
            throw new Error(
                `Columns in table '${TableNameTransactions}' do not match expected transaction` +
                    ` header structure. Try deleting the entire table.`
            );
        }

        // Load the values from the table so that empty rows can be found and deleted:
        const tranTableBodyRange = tranTable.getDataBodyRange();
        tranTable.rows.load(["count"]);
        tranTableBodyRange.load("values");
        await context.excel.sync();

        const tranTableRowCount = tranTable.rows.count;
        const tranTableValues = tranTableBodyRange.values;

        // Delete empty rows (start from bottom to avoid index shift):
        let countEmptyRowsDeleted = 0;
        for (let r = tranTableValues.length - 1; r >= 0; r--) {
            // If we are at top row and all rows so far were deleted, skip.
            // This is because tables may never have zero rows.
            if (r === 0 && countEmptyRowsDeleted === tranTableRowCount - 1) {
                continue;
            }
            const tranTableValueRow = tranTableValues[r]!;
            const isRowEmpty = tranTableValueRow.every((val) => isNullOrWhitespace(val));
            if (isRowEmpty) {
                tranTable.rows.getItemAt(r).delete();
                countEmptyRowsDeleted++;
            }
        }
        await context.excel.sync();
        console.debug(`Deleted ${countEmptyRowsDeleted} empty rows from table '${tranTable.name}'.`);

        // Refresh the table data after the earlier deletion:
        tranTable.rows.load(["count", "items"]);
        tranTable.columns.load(["count", "items"]);
        await context.excel.sync();
        for (const col of tranTable.columns.items) {
            col.load("name");
        }
        await context.excel.sync();

        // Load data from the table into `existingTrans`:
        const tranTableColNames = tranTable.columns.items.map((col) => col.name.trim());

        // Apply formatting:
        for (const tabCol of tranTable.columns.items) {
            if (tabCol.name === LunchIdColumnName) {
                const tabColRange = tabCol.getDataBodyRange();
                tabColRange.format.font.size = 6;
                tabColRange.format.verticalAlignment = "Center";
                tabColRange.format.horizontalAlignment = "Right";
            }

            const format = tranColumnsSpecs.getByKey(tabCol.name)?.format;
            if (!format) {
                continue;
            }

            const tabColRange = tabCol.getDataBodyRange();
            tabColRange.numberFormat = [[format]];
        }

        await context.excel.sync();

        const existingTrans = new IndexedMap<number, TransactionRowData>();

        for (let r = 0; r < tranTable.rows.count; r++) {
            // Load the values of this row:
            const rowRange = tranTable.rows.getItemAt(r).getRange();
            rowRange.load(["values", "address"]);
            await context.excel.sync();
            const rowRangeValues = rowRange.values[0]!;

            // Build the data object:
            let isEmptyRow = true;
            const rowDataValues: Record<string, TransactionRowValue> = {};
            for (let c = 0; c < rowRangeValues.length; c++) {
                const colName = tranTableColNames[c]!;
                if (rowRangeValues[c] === undefined) {
                    throw new Error(`Column #${c} ('${colName}') not specified for item on row ${r}.`);
                }
                if (rowRangeValues[c] !== "") {
                    isEmptyRow = false;
                }
                rowDataValues[colName] = rowRangeValues[c];
            }

            if (!isEmptyRow) {
                // Add the object to the loaded collection by ID and by order:
                const lunchIdStr = rowDataValues[LunchIdColumnName];
                if (!lunchIdStr) {
                    throw new Error(`${LunchIdColumnName} not specified for item on row ${r}.`);
                }
                const lunchId = Number(lunchIdStr);
                if (!Number.isInteger(lunchId)) {
                    throw new Error(`Invalid ${LunchIdColumnName}-value ('${lunchIdStr}') for item on row ${r}.`);
                }

                const rowInfo = { values: rowDataValues, range: rowRange };
                existingTrans.tryAdd(lunchId, rowInfo);
            }
        }

        console.debug(`Read ${existingTrans.length} existing data rows from table '${tranTable.name}'.`);

        // Fetch transactions:

        // Fetch Transactions from the Cloud:
        //

        console.log("Will now fetch transactions from Lunch Money.");

        const startUtcDateStr = formatDateUtc(startDate);
        const endUtcDateStr = formatDateUtc(endDate);
        const fetchedResponseText = await authorizedFetch(
            "GET",
            `transactions?start_date=${startUtcDateStr}&end_date=${endUtcDateStr}`,
            `get all transactions between ${startUtcDateStr} and ${endUtcDateStr} UTC`
        );

        // Parse fetched Transactions:
        const fetched: { transactions: Lunch.Transaction[]; has_more: boolean } = JSON.parse(fetchedResponseText);
        console.log("Transactions fetched. Length:", fetched.transactions.length, "Has_more:", fetched.has_more);
        //console.debug("Fetched transactions:", fetched.transactions);

        const receivedTrans = new IndexedMap<number, Transaction>();

        // Parsed Plaid data for each transaction:
        let countPlaidMetadataObjectsParsed = 0;
        for (let t = 0; t < fetched.transactions.length; t++) {
            const fetchedTran = fetched.transactions[t];
            if (!fetchedTran) {
                continue;
            }

            const fetchedTranId = fetchedTran.id;
            if (!Number.isInteger(fetchedTranId)) {
                throw new Error(
                    `Cannot parse ID of fetched transaction #${t}. Integer ID expected. (Actual id='${fetchedTranId}'.)`
                );
            }

            const tran: Transaction = {
                trn: fetchedTran,
                pld: null,
                tag: new Map<string, Set<string>>(),
                grpMoniker: "",
                id: fetchedTranId,
            };

            receivedTrans.tryAdd(tran.id, tran);

            const plaidDataStr = fetchedTran.plaid_metadata;
            if (typeof plaidDataStr === "string") {
                try {
                    const plaidMetadata: Lunch.PlaidMetadata = JSON.parse(plaidDataStr);
                    tran.pld = plaidMetadata;
                    countPlaidMetadataObjectsParsed++;
                } catch (err) {
                    console.error(
                        `Cannot parse plaid_metadata for fetched transaction #${t}.`,
                        "plaidDataStr:",
                        plaidDataStr
                    );
                    throw new Error(
                        `Cannot parse plaid_metadata for fetched transaction #${t} (id=${fetchedTran?.id}).`,
                        {
                            cause: err,
                        }
                    );
                }
            }
        }
        console.log(
            `Plaid Metadata objects parsed: ${countPlaidMetadataObjectsParsed} / ${fetched.transactions.length}.` +
                "\n(Not all transactions will have Plaid Metadata. E.g., groups, split transactions, transactions" +
                " imported form sources other than Plaid.)"
        );

        // Parse Tags for all transactions:

        const allReceivedTags = new Map<string, Set<string>>();

        for (let t = 0; t < receivedTrans.length; t++) {
            const tran = receivedTrans.getByIndex(t);
            if (!tran || !tran.trn.tags) {
                continue;
            }

            // For each tag of this transaction:
            const tags: { name: string; id: number }[] = tran.trn.tags;
            for (const tag of tags) {
                // Parse the tag:
                const tagInfo = parseTag(tag.name);
                //console.debug(`Transaction '${tran.trn.id}'. tagInfo:`, tagInfo);

                // Add the tag to the list of this transaction:
                if (!tran.tag.has(tagInfo.group)) {
                    tran.tag.set(tagInfo.group, new Set<string>());
                }
                tran.tag.get(tagInfo.group)!.add(tagInfo.value);

                // Add the tag to the list of all received tags:
                if (!allReceivedTags.has(tagInfo.group)) {
                    allReceivedTags.set(tagInfo.group, new Set<string>());
                }
                allReceivedTags.get(tagInfo.group)!.add(tagInfo.value);
            }
        }

        console.log(`Received transactions contains tags from ${allReceivedTags.size} different groups.`);

        // To do: Resolve group names.

        // Create rows to add:

        // Todo: deal with already existing data.

        const tranRowsToAdd: (string | boolean | number)[][] = [];
        let countExistingTransDetected = 0;
        for (const tran of receivedTrans) {
            if (existingTrans.has(tran.id)) {
                countExistingTransDetected++;
                continue;
            }

            const rowToAdd: (string | boolean | number)[] = [];
            for (const colName of tranTableColNames) {
                rowToAdd.push(getTransactionColumnValue(tran, colName, tranColumnsSpecs));
            }
            tranRowsToAdd.push(rowToAdd);
        }

        console.log(
            `${receivedTrans.length} received transactions include ${tranRowsToAdd.length} new items` +
                ` and ${countExistingTransDetected} existing items.`
        );

        tranTable.rows.load(["items", "count"]);
        if (tranRowsToAdd.length > 0) {
            tranTable.rows.add(0, tranRowsToAdd);
            tranTable.getRange().format.autofitColumns();
        }

        await context.excel.sync();
    } catch (err) {
        console.error(err);
        errorMsgCell.values = [[`ERR: ${errorTypeMessageString(err)}`]];
        errorMsgCell.format.font.color = "#FF0000";
        await context.excel.sync();
        throw err;
    }
}
