/// <reference types="office-js" />

import { ErrorAggregator } from "src/util/ErrorAggregator";

export async function syncBlock(excelContext: Excel.RequestContext, block: () => void | Promise<void>): Promise<void> {
    const errors = new ErrorAggregator();

    try {
        await block();
    } catch (err) {
        errors.add(err);
    }

    try {
        await excelContext.sync();
    } catch (err) {
        errors.add(err);
    }

    errors.throwIfHasErrors();
}

export async function findTableByName(
    tableName: string,
    excelContext: Excel.RequestContext
): Promise<{ table: Excel.Table; sheet: Excel.Worksheet; range: Excel.Range } | null> {
    // Get the table by name (returns null object if not found)
    const table = excelContext.workbook.tables.getItemOrNullObject(tableName);

    table.load("name, id, worksheet, isNullObject");
    await excelContext.sync();

    // Check if the table exists
    if (table.isNullObject) {
        return null;
    }

    const sheet = table.worksheet;
    const range = table.getRange();

    // These properties are likely required after the func returns
    sheet.load("name, id");
    range.load("address");
    await excelContext.sync();

    return { table, sheet, range };
}

/** Finds a table by name. If the table does not exist at all, the returned promise resolves to null.
 * However, if the table exists, but NOT on the specified sheet, the an error is thrown. */
export async function findTableByNameOnSheet(
    tableName: string,
    expectedSheet: Excel.Worksheet,
    excelContext: Excel.RequestContext
): Promise<{ table: Excel.Table; sheet: Excel.Worksheet; range: Excel.Range } | null> {
    // Load target sheet id:
    expectedSheet.load(["id"]);
    await excelContext.sync();
    const expectedSheetId = expectedSheet.id;

    // No table => return null:
    const existingTableInfo = await findTableByName(tableName, excelContext);
    if (existingTableInfo === null) {
        return null;
    }

    // There is an existing tags table, but it is not on this sheet:
    if (existingTableInfo.sheet.id !== expectedSheetId) {
        throw new Error(
            `Table '${tableName}' exists on the wrong sheet ('${existingTableInfo.range.address}').` +
                "\n NOTE: Don't name any objects using prefix 'LM.';" +
                " Don't edit any auto-created sheets, except where specified." +
                "\n SOLUTION: Back-up your data, delete or rename the unexpected table, reload the Add-In."
        );
    }

    return existingTableInfo;
}

export async function ensureSheetActive(
    sheetName: string,
    excelContext: Excel.RequestContext
): Promise<Excel.Worksheet> {
    let sheet = excelContext.workbook.worksheets.getItemOrNullObject(sheetName);
    sheet.load("name, id, isNullObject");
    await excelContext.sync();

    if (sheet.isNullObject) {
        sheet = excelContext.workbook.worksheets.add(sheetName);
        await excelContext.sync();
    }

    sheet.activate();
    sheet.load("name, id, isNullObject");
    await excelContext.sync();

    return sheet;
}

export function timeStrToExcel(datetimeStr: string): number {
    const dt = new Date(datetimeStr);

    if (isNaN(dt.getTime())) {
        throw new Error(`Invalid date string '${datetimeStr}'.`);
    }

    // Excel epoch (1900 system): 1899-12-30
    const excelEpoch = Math.abs(Date.UTC(1899, 11, 30));

    const msecPerDay = 24 * 60 * 60 * 1000;

    // Convert JS ms → Excel days
    return (dt.getTime() + excelEpoch) / msecPerDay;
}

export function getRangeBasedOn(
    sheet: Excel.Worksheet,
    base: { row: number; col: number },
    startRowOffsFromBase: number,
    startColumnOffsFromBase: number,
    rowCount: number,
    columnCount: number
): Excel.Range {
    return sheet.getRangeByIndexes(
        base.row + startRowOffsFromBase,
        base.col + startColumnOffsFromBase,
        rowCount,
        columnCount
    );
}

export function parseOnSheetAddress(address: string): string {
    const sepPos = address.indexOf("!");
    return sepPos < 0 ? address.trim() : address.substring(sepPos + 1).trim();
}
