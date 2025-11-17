/// <reference types="office-js" />

export async function findTableByName(
    tableName: string,
    excelContext: Excel.RequestContext
): Promise<{ table: Excel.Table; sheet: Excel.Worksheet; range: Excel.Range } | null> {
    // Get the table by name (returns null object if not found)
    const table = excelContext.workbook.tables.getItemOrNullObject(tableName);

    if (!table) {
        return null;
    }

    const sheet = table.worksheet;
    const range = table.getRange();

    // Load the properties we need
    table.load("name, id, isNullObject");
    sheet.load("name, id");
    range.load("address");

    // Sync to execute the queued commands
    await excelContext.sync();

    // Check if the table exists
    if (table.isNullObject) {
        return null;
    }

    return { table, sheet, range };
}
