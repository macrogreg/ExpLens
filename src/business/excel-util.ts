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

export async function ensureSheetActive(sheetName: string, excelContext: Excel.RequestContext): Promise<Excel.Worksheet> {
    let sheet = excelContext.workbook.worksheets.getItemOrNullObject(sheetName);
    sheet.load("name, id, isNullObject");
    await excelContext.sync();

    if (sheet.isNullObject) {
        sheet = excelContext.workbook.worksheets.add(sheetName);
        await excelContext.sync();
    }

    sheet.activate();
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
