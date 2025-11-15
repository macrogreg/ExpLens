/// <reference types="office-js" />
// Business logic for LunchMoney Excel Add-In

/**
 * Downloads transactions and writes parameters to the top cells of the current Excel worksheet.
 * @param apiToken LunchMoney API token
 * @param fromDate ISO string for FROM date
 * @param toDate ISO string for TO date
 * @throws Error if Excel JS APIs are not available
 */
export async function downloadTransactions(apiToken: string, fromDate: string, toDate: string): Promise<void> {
    if (typeof Excel === "undefined" || typeof Excel.run !== "function") {
        throw new Error("Excel JS APIs are not available. This function must be run inside Excel.");
    }

    await Excel.run(async (context: Excel.RequestContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getRange("A1").values = [[apiToken]];
        sheet.getRange("A2").values = [[fromDate]];
        sheet.getRange("A3").values = [[toDate]];
        await context.sync();
    });
}
