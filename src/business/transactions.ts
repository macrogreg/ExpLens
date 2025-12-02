/// <reference types="office-js" />

import { errorTypeMessageString, formatDateUtc } from "src/util/format_util";
import { datetimeToExcel, findTableByNameOnSheet, getRangeBasedOn } from "./excel-util";
import type { Transaction, TransactionColumnSpec, TransactionRowData, TransactionRowValue } from "./transaction-tools";
import {
    getTransactionColumnValue,
    SpecialColumnNames,
    createTransactionColumnsSpecs,
    tryGetTagGroupFromColumnName,
    getTagColumnsPosition,
    formatTagGroupColumnHeader,
} from "./transaction-tools";
import { isNullOrWhitespace } from "src/util/string_util";
import type { TagValuesCollection } from "./tags";
import { parseTag } from "./tags";
import { authorizedFetch } from "./fetch-tools";
import type * as Lunch from "./lunchmoney-types";
import { IndexedMap } from "./IndexedMap";
import type { SyncContext } from "./sync-driver";
import { useSheetProgressTracker } from "src/composables/sheet-progress-tracker";

export const SheetNameTransactions = "EL.Transactions";
const TableNameTransactions = "EL.TransactionsTable";

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

// If the sync context specified columns that were not already present in the table, insert them:
async function insertMissingTagColumns(tranTable: Excel.Table, context: SyncContext) {
    tranTable.columns.load(["count", "items"]);
    tranTable.rows.load(["count"]);
    await context.excel.sync();

    // Set of all required tag columns:
    const reqTagColNames = context.tags.assignable.keys().map((gn) => formatTagGroupColumnHeader(gn));
    const missingColNames = new Set(reqTagColNames);

    // Scroll through table, remove all encountered tag columns from search set:
    let firstTagColNum: number | undefined = undefined;
    for (let c = 0; c < tranTable.columns.count; c++) {
        const col = tranTable.columns.getItemAt(c).load(["name"]);
        await context.excel.sync();

        // If no more outstanding column names, be done:
        missingColNames.delete(col.name);
        if (missingColNames.size === 0) {
            return;
        }

        if (tryGetTagGroupFromColumnName(col.name)) {
            firstTagColNum = firstTagColNum === undefined ? c : firstTagColNum;
        }
    }

    if (firstTagColNum === undefined) {
        firstTagColNum = getTagColumnsPosition();
    }

    //Add required columns:
    const colNamesToAdd = [...missingColNames].sort().reverse();
    for (const cn of colNamesToAdd) {
        tranTable.columns.add(firstTagColNum, undefined, cn);
    }

    await context.excel.sync();
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

async function applyColumnFormatting(
    tranTable: Excel.Table,
    tranColumnsSpecs: IndexedMap<string, TransactionColumnSpec>,
    context: SyncContext
) {
    console.debug("Will apply formatting to table columns...");
    const msStartApplyFormatting = performance.now();

    for (const tabCol of tranTable.columns.items) {
        const colName = tabCol.name;
        const colSpec = tranColumnsSpecs.getByKey(colName);
        if (colSpec === undefined) {
            continue;
        }

        const tabColRange = tabCol.getDataBodyRange();

        const numFormat = colSpec.numberFormat;
        if (numFormat) {
            tabColRange.numberFormat = [[numFormat]];
        }

        const formatFn = colSpec.formatFn;
        if (formatFn) {
            try {
                const formatFnRes = formatFn(tabColRange.format, tabColRange.dataValidation, context);
                await formatFnRes;
                await context.excel.sync();
            } catch (err) {
                console.error(
                    `The formatFn of the column spec for '${colName}' threw an error.` +
                        `\n    '${errorTypeMessageString(err)}'` +
                        `\n\n We will skip over this and continue, but this needs to be corrected.`,
                    `\n\nERR:\n`,
                    err
                );
            }
        }
    }

    await context.excel.sync();

    console.debug(
        `Formatting applied to table columns.\n    Time taken: ${performance.now() - msStartApplyFormatting} msec.`
    );
}

function setEditableHintRangeFormat(range: Excel.Range, editableState: "Read-Only" | "Editable") {
    range.clear();
    range.format.horizontalAlignment = Excel.HorizontalAlignment.center;
    range.format.verticalAlignment = Excel.VerticalAlignment.center;
    range.format.font.size = 10;

    switch (editableState) {
        case "Read-Only":
            range.format.fill.color = "#f2ceef";
            range.format.font.color = "#d76dcc";
            break;
        case "Editable":
            range.format.fill.color = "#b5e6a2";
            range.format.font.color = "#4ea72e";
            break;
    }

    range.dataValidation.prompt = {
        showPrompt: true,
        title: `${editableState} column`,
        message:
            editableState === "Read-Only"
                ? "Let ExpLens manage it for you."
                : editableState === "Editable"
                  ? "Select a value from the dropdown"
                  : "",
    };
}

async function printSheetHeaders(context: SyncContext) {
    context.sheets.trans.getRange("B2").values = [["Transactions"]];
    context.sheets.trans.getRange("B2:E2").style = "Heading 1";

    const tabRdOnlyMsgRange = context.sheets.trans.getRange("B3:E3");
    tabRdOnlyMsgRange.clear();
    tabRdOnlyMsgRange.merge();
    tabRdOnlyMsgRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;
    tabRdOnlyMsgRange.format.verticalAlignment = Excel.VerticalAlignment.center;
    tabRdOnlyMsgRange.format.fill.color = "#fff8dc";
    tabRdOnlyMsgRange.format.font.color = "d76dcc";
    tabRdOnlyMsgRange.format.font.size = 10;

    tabRdOnlyMsgRange.getCell(0, 0).values = [["This tab is managed by ExpLens. Only modify specific columns:"]];

    const tabRwAreasDocRange = context.sheets.trans.getRange("B4:B4");
    setEditableHintRangeFormat(tabRwAreasDocRange, "Editable");
    tabRwAreasDocRange.format.font.bold = true;
    tabRwAreasDocRange.values = [["Editable Columns"]];

    const tabRoAreasDocRange = context.sheets.trans.getRange("C4:C4");
    setEditableHintRangeFormat(tabRoAreasDocRange, "Read-Only");
    tabRoAreasDocRange.format.font.bold = true;
    tabRoAreasDocRange.values = [["Read-Only Columns"]];

    await context.excel.sync();
}

async function createEditableHintHeader(tranTable: Excel.Table, context: SyncContext) {
    tranTable.columns.load(["count"]);
    const tranTableRange = tranTable.getRange();
    tranTableRange.load(["address", "rowIndex", "columnIndex"]);
    await context.excel.sync();

    const hintRowOffs = { row: tranTableRange.rowIndex - 1, col: tranTableRange.columnIndex };

    // Create Read-Only vs Editable marker row:
    const editableColumnNames = new Set<string>(["Category"]);
    for (const gn of context.tags.assignable.keys()) {
        editableColumnNames.add(formatTagGroupColumnHeader(gn));
    }

    for (let c = 0; c < tranTable.columns.count; c++) {
        const tabCol = tranTable.columns.getItemAt(c);
        tabCol.load(["name"]);
        await context.excel.sync();

        const colKind = editableColumnNames.has(tabCol.name) ? "Editable" : "Read-Only";
        const colHintCell = getRangeBasedOn(context.sheets.trans, hintRowOffs, 0, c, 1, 1);
        setEditableHintRangeFormat(colHintCell, colKind);
    }
}

async function createInfoRow(tranTable: Excel.Table, context: SyncContext) {
    const tranTableRange = tranTable.getRange();
    tranTableRange.load(["address", "rowIndex", "columnIndex", "name"]);
    await context.excel.sync();

    const infoRowOffs = { row: tranTableRange.rowIndex - 2, col: tranTableRange.columnIndex };

    // Count:
    const countTransLabelRange = getRangeBasedOn(context.sheets.trans, infoRowOffs, 0, 0, 1, 1);
    countTransLabelRange.format.fill.clear();
    countTransLabelRange.format.font.color = "#7e350e";
    countTransLabelRange.format.font.bold = true;
    countTransLabelRange.format.horizontalAlignment = "Right";
    countTransLabelRange.values = [["Count:"]];

    const countTransFormulaRange = getRangeBasedOn(context.sheets.trans, infoRowOffs, 0, 1, 1, 1);
    countTransFormulaRange.format.fill.color = "#f2f2f2";
    countTransFormulaRange.format.font.color = "#7e350e";
    countTransFormulaRange.format.font.bold = true;
    countTransFormulaRange.format.horizontalAlignment = "Left";
    countTransFormulaRange.formulas = [[`="  " & COUNTA(${tranTable.name}[${SpecialColumnNames.LunchId}])`]];

    // Last successful sync:
    const lastCompletedSyncLabelRange = getRangeBasedOn(context.sheets.trans, infoRowOffs, 0, 3, 1, 1);
    lastCompletedSyncLabelRange.format.fill.clear();
    lastCompletedSyncLabelRange.format.font.color = "#7e350e";
    lastCompletedSyncLabelRange.format.font.bold = true;
    lastCompletedSyncLabelRange.format.horizontalAlignment = "Right";
    lastCompletedSyncLabelRange.values = [["Last completed download data version / time:"]];

    const lastCompletedSyncVersionRange = getRangeBasedOn(context.sheets.trans, infoRowOffs, 0, 4, 1, 1);
    lastCompletedSyncVersionRange.format.fill.color = "#f2f2f2";
    lastCompletedSyncVersionRange.format.font.color = "#7e350e";
    lastCompletedSyncVersionRange.format.font.bold = true;
    lastCompletedSyncVersionRange.format.horizontalAlignment = "Left";
    lastCompletedSyncVersionRange.formulas = [[`="  " & "${context.currentSync.version}"`]];

    const lastCompletedSyncTimeRange = getRangeBasedOn(context.sheets.trans, infoRowOffs, 0, 5, 1, 1);
    lastCompletedSyncTimeRange.format.fill.color = "#f2f2f2";
    lastCompletedSyncTimeRange.format.font.color = "#7e350e";
    lastCompletedSyncTimeRange.format.font.bold = true;
    lastCompletedSyncTimeRange.format.horizontalAlignment = "Left";
    lastCompletedSyncTimeRange.values = [[datetimeToExcel(context.currentSync.utc, true)]];
    lastCompletedSyncTimeRange.numberFormat = [["  yyyy-mm-dd  HH:mm:ss"]];

    console.log(`Current time 1: '${String(context.currentSync.utc)}'`);
    console.log(`Current time 2: '${String(context.currentSync.utc.toLocaleString())}'`);
    console.log(`Current time 3: '${String(context.currentSync.utc.toUTCString())}'`);
    console.log(`Current time 3: '${String(context.currentSync.utc.getTime())}'`);

    await context.excel.sync();
}

export async function downloadTransactions(startDate: Date, endDate: Date, context: SyncContext) {
    const transSheetProgressTracker = useSheetProgressTracker(31, 90, context);
    transSheetProgressTracker.setPercentage(0);

    // Activate the sheet:
    context.sheets.trans.activate();
    await context.excel.sync();

    // Clear and prepare the location for printing potential errors:
    const errorMsgBackgroundRange = context.sheets.trans.getRange("B5:F5");
    errorMsgBackgroundRange.clear();
    // errorMsgBackgroundRange.merge();
    // errorMsgBackgroundRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;
    const errorMsgRange = errorMsgBackgroundRange.getCell(0, 0);
    await context.excel.sync();

    try {
        transSheetProgressTracker.setPercentage(5);
        await printSheetHeaders(context);

        const tranColumnsSpecs: IndexedMap<string, TransactionColumnSpec> = createTransactionColumnsSpecs(context);

        // Make first (empty) column slim:
        context.sheets.trans.getRange("A:A").format.columnWidth = 15;
        await context.excel.sync();

        // if ("Testing errors".length < 100) {
        //     throw new Error("A test error was thrown. A detailed description of this error should be displayed.");
        // }

        // Is there an existing Transactions table?
        const prevTranTableInfo = await findTableByNameOnSheet(
            TableNameTransactions,
            context.sheets.trans,
            context.excel
        );

        // If there is no existing table, create an empty one:
        const tranTable =
            prevTranTableInfo === null
                ? await createNewTranTable(tranColumnsSpecs, context.sheets.trans, context.excel)
                : prevTranTableInfo.table;

        // If the sync context specified columns that were not already present in the table, insert them:
        await insertMissingTagColumns(tranTable, context);

        // Create the RO/RW hints header above the table:
        await createEditableHintHeader(tranTable, context);

        // Freeze table head:
        {
            context.sheets.trans.freezePanes.unfreeze();
            await context.excel.sync();

            const tranTableHeaderRange = tranTable.getHeaderRowRange();
            tranTableHeaderRange.load(["rowIndex"]);
            await context.excel.sync();

            context.sheets.trans.freezePanes.freezeRows(tranTableHeaderRange.rowIndex + 1);
            await context.excel.sync();
        }

        // Load the column names actually present in the table:
        tranTable.columns.load(["count", "items"]);
        await context.excel.sync();
        for (const col of tranTable.columns.items) {
            col.load("name");
        }
        await context.excel.sync();

        // Cache the actually present Transactions table Column Names:
        const tranTableColNames: string[] = tranTable.columns.items.map((col) => col.name.trim());

        // Ensure the ID column exists:
        if (!tranTableColNames.includes(SpecialColumnNames.LunchId)) {
            throw new Error(
                `Table '${tranTable.name}' (${prevTranTableInfo === null ? "newly created" : "pre-existing"})` +
                    ` does not contain the expected column '${SpecialColumnNames.LunchId}'.`
            );
        }

        // Validate that the actual column names match the spec (this ignores Tag columns):
        if (!isColumnNamingEquivalent(tranColumnsSpecs, tranTableColNames)) {
            throw new Error(
                `Columns in table '${TableNameTransactions}' do not match expected transaction` +
                    ` header structure. Try deleting the entire table.`
            );
        }

        transSheetProgressTracker.setPercentage(10);

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
        tranTable.rows.load(["count", "items"]);
        await context.excel.sync();
        console.debug(`Deleted ${countEmptyRowsDeleted} empty rows from table '${tranTable.name}'.`);

        transSheetProgressTracker.setPercentage(15);

        // Load data from the table into `existingTrans`:

        console.debug(`Will read ${tranTable.rows.count} existing data rows from '${tranTable.name}'...`);
        const msStartReadPreexistingData = performance.now();

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
                const lunchIdStr = rowDataValues[SpecialColumnNames.LunchId];
                if (!lunchIdStr) {
                    throw new Error(`${SpecialColumnNames.LunchId} not specified for item on row ${r}.`);
                }
                const lunchId = Number(lunchIdStr);
                if (!Number.isInteger(lunchId)) {
                    throw new Error(
                        `Invalid ${SpecialColumnNames.LunchId}-value ('${lunchIdStr}') for item on row ${r}.`
                    );
                }

                const rowInfo = { values: rowDataValues, range: rowRange };
                existingTrans.tryAdd(lunchId, rowInfo);
            }
        }

        console.debug(
            `Done reading ${existingTrans.length} existing data rows from table '${tranTable.name}'.` +
                `\n    Time taken: ${performance.now() - msStartReadPreexistingData} msec.`
        );

        transSheetProgressTracker.setPercentage(30);

        // Fetch transactions:

        console.log("Will fetch transactions from Lunch Money...");
        const msStartFetchTransactions = performance.now();

        const startUtcDateStr = formatDateUtc(startDate);
        const endUtcDateStr = formatDateUtc(endDate);
        const fetchedResponseText = await authorizedFetch(
            "GET",
            `transactions?start_date=${startUtcDateStr}&end_date=${endUtcDateStr}`,
            `get all transactions between ${startUtcDateStr} and ${endUtcDateStr} UTC`
        );

        console.log(`Transactions fetched.\n    Time taken: ${performance.now() - msStartFetchTransactions} msec.`);

        transSheetProgressTracker.setPercentage(45);

        // Parse fetched Transactions:
        const fetched: { transactions: Lunch.Transaction[]; has_more: boolean } = JSON.parse(fetchedResponseText);
        console.log("Transactions parsed. Length:", fetched.transactions.length, "Has_more:", fetched.has_more);
        //console.debug("Fetched transactions:", fetched.transactions);

        if (fetched.has_more) {
            console.error("There are more transactions to fetch, but this is not yet supported!");
        }

        transSheetProgressTracker.setPercentage(50);

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

        transSheetProgressTracker.setPercentage(55);

        // Parse Tags for all received transactions:

        const allReceivedTags: TagValuesCollection = new Map<string, Set<string>>();

        for (let t = 0; t < receivedTrans.length; t++) {
            const tran = receivedTrans.getByIndex(t);
            if (!tran || !tran.trn.tags) {
                continue;
            }

            // For each tag of this transaction, and add it to the tag values collections:
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

        transSheetProgressTracker.setPercentage(60);

        // Go over downloaded transactions and decide what to do with Each:
        // - Existing transitions:
        //   - In No-Update mode: Just Skip;
        //   - In Update mode:
        //     Loop over each property of Existing, compare with Received; If different - track for update;
        //     If any differences found: Update table right away;
        // - New transactions:
        //   Create new data row for insertion;
        // Finally, insert all newly create rows into the table.

        const colIndexLastSyncVersion = tranTableColNames.findIndex((cn) => cn === SpecialColumnNames.LastSyncVersion);

        const tranRowsToAdd: (string | boolean | number)[][] = [];
        const countExistingTransDetected = {
            sameAsReceived: 0,
            differentFromReceived: 0,
            notComparedWithReceived: 0,
        };

        for (const tran of receivedTrans) {
            // If Transaction already in table (existing):
            const exTran = existingTrans.getByKey(tran.id);
            if (exTran !== undefined) {
                // If replacing existing transitions is NOT required, just skip it:
                if (!context.isReplaceExistingTransactions) {
                    countExistingTransDetected.notComparedWithReceived++;
                    continue;
                }

                // Replacing the existing transitions is IS required. So:
                // Loop over the existing data and compare it with received data:
                exTran.range.load(["formulas"]);
                await context.excel.sync();
                const tranDataVals = exTran.range.values[0]!;
                const tranFormulas = exTran.range.formulas[0]!;
                let needsUpdating = false;
                for (let c = 0; c < tranDataVals.length; c++) {
                    // Skip the sync version column, as it is always different:
                    if (c === colIndexLastSyncVersion) {
                        continue;
                    }

                    // Check whether received data is different, and if so - track for update:
                    const colName = tranTableColNames[c]!;
                    const existingColVal = tranDataVals[c];
                    const receivedColVal = getTransactionColumnValue(tran, colName, tranColumnsSpecs);
                    if (existingColVal !== receivedColVal) {
                        if (
                            typeof existingColVal === "number" &&
                            typeof receivedColVal === "string" &&
                            Number(receivedColVal.trim()) === existingColVal
                        ) {
                            // Excel eagerly converts strings to numbers.
                            // If after that conversions, values match, we consider them Equal.
                        } else if (tranFormulas[c] === receivedColVal) {
                            // The existing value doesn't match, but the formula does.
                            // No need to update, but we still copy the formula
                            // in case we detect the need to update based on another column:
                            tranDataVals[c] = tranFormulas[c];
                        } else {
                            // NO match for either column or formula. Track for UPDATE:
                            tranDataVals[c] = receivedColVal;
                            needsUpdating = true;
                        }
                    }
                }

                // If the received data is different, update the transaction row:
                if (needsUpdating) {
                    tranDataVals[colIndexLastSyncVersion] = context.currentSync.version;
                    exTran.range.values = [tranDataVals];
                    countExistingTransDetected.differentFromReceived++;
                    await context.excel.sync();
                } else {
                    countExistingTransDetected.sameAsReceived++;
                }
            } else {
                // The `exTran` is undefined, i.e. received `tran` is new.
                // Initialize a new data row based on the received transaction:
                const rowToAdd: (string | boolean | number)[] = [];

                for (const colName of tranTableColNames) {
                    if (colName === SpecialColumnNames.LastSyncVersion) {
                        rowToAdd.push(context.currentSync.version);
                    } else {
                        rowToAdd.push(getTransactionColumnValue(tran, colName, tranColumnsSpecs));
                    }
                }

                // Add the transaction data row to the list of rows to add:
                tranRowsToAdd.push(rowToAdd);
            }
        }

        console.log(
            `${receivedTrans.length} received transactions were processed.`,
            "context.isReplaceExistingTransactions: ",
            context.isReplaceExistingTransactions,
            "countExistingTransDetected: ",
            countExistingTransDetected,
            "tranRowsToAdd.length: ",
            tranRowsToAdd.length
        );

        transSheetProgressTracker.setPercentage(67);

        // Insert new transaction rows:
        tranTable.rows.load(["items", "count"]);
        if (tranRowsToAdd.length > 0) {
            console.debug(`Inserting ${tranRowsToAdd.length} rows into the table...`);
            const msStartAddTableRows = performance.now();

            tranTable.rows.add(0, tranRowsToAdd);
            await context.excel.sync();

            console.log(`Rows inserted.\n    Time taken: ${performance.now() - msStartAddTableRows} msec.`);
        }

        transSheetProgressTracker.setPercentage(75);

        // Sort the table:

        const sortFields: Excel.SortField[] = [
            {
                key: tranTableColNames.findIndex((cn) => cn === "date"),
                sortOn: Excel.SortOn.value,
                ascending: false,
            },
            {
                key: tranTableColNames.findIndex((cn) => cn === "plaid:authorized_datetime"),
                sortOn: Excel.SortOn.value,
                ascending: false,
            },
            {
                key: tranTableColNames.findIndex((cn) => cn === "Account"),
                sortOn: Excel.SortOn.value,
                ascending: true,
            },
            {
                key: tranTableColNames.findIndex((cn) => cn === "Payer"),
                sortOn: Excel.SortOn.value,
                ascending: true,
            },
            {
                key: tranTableColNames.findIndex((cn) => cn === "payee"),
                sortOn: Excel.SortOn.value,
                ascending: true,
            },
            {
                key: tranTableColNames.findIndex((cn) => cn === "plaid:datetime"),
                sortOn: Excel.SortOn.value,
                ascending: false,
            },
            {
                key: tranTableColNames.findIndex((cn) => cn === "SpecialColumnNames.LunchId"),
                sortOn: Excel.SortOn.value,
                ascending: true,
            },
        ].filter((f) => f.key >= 0);

        tranTable.sort.apply(sortFields);
        await context.excel.sync();

        transSheetProgressTracker.setPercentage(80);

        // Apply formatting to all columns and rows:
        await applyColumnFormatting(tranTable, tranColumnsSpecs, context);

        transSheetProgressTracker.setPercentage(90);

        // Add info on transactions count, and on version and time of the last successful sync:
        await createInfoRow(tranTable, context);

        // Auto-fit the table:
        tranTable.getRange().format.autofitColumns();
        await context.excel.sync();

        transSheetProgressTracker.setPercentage(100);
    } catch (err) {
        console.error(err);
        errorMsgRange.values = [[`ERR: ${errorTypeMessageString(err)}`]];
        errorMsgRange.format.font.color = "#FF0000";
        await context.excel.sync();
        throw err;
    }
}
