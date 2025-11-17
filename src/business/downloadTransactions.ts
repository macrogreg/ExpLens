/// <reference types="office-js" />

import { isNotNullOrWhitespaceStr } from "src/util/string_util";
import { useApiToken } from "./apiToken";

// Business logic for LunchMoney Excel Add-In

/**
 * Downloads transactions and writes parameters to the top cells of the current Excel worksheet.
 * @param apiToken LunchMoney API token
 * @param fromDate ISO string for FROM date
 * @param toDate ISO string for TO date
 * @throws Error if Excel JS APIs are not available
 */
export async function downloadTransactions(fromDate: string, toDate: string): Promise<void> {
    if (typeof Excel === "undefined" || typeof Excel.run !== "function") {
        throw new Error("Excel JS APIs are not available. This function must be run inside Excel.");
    }

    const apiToken = useApiToken().value();
    if (!isNotNullOrWhitespaceStr(apiToken)) {
        console.log("Will not try to download transactions, because no API Token is set.");
        return;
    }

    await Excel.run(async (context: Excel.RequestContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.getRange("A1").values = [[apiToken]];
        sheet.getRange("A2").values = [[fromDate]];
        sheet.getRange("A3").values = [[toDate]];

        await context.sync();

        // Fetch transactions
        const headers = new Headers();
        headers.append("Authorization", `Bearer ${apiToken}`);
        const response = await fetch("https://dev.lunchmoney.app/v1/transactions", {
            method: "GET",
            headers,
            redirect: "follow",
        });
        const result: { transactions: Record<string, unknown>[]; has_more: boolean } = await response.json();

        console.log("HEADERS:", headers);
        console.log("RESPONSE:", result);

        // Write transaction count to B5
        sheet.getRange("B5").values = [[result.transactions.length]];

        // If there are transactions, write them starting at B6
        if (result.transactions.length > 0) {
            // Get all keys from the first transaction for column headers
            const firstTrans = result.transactions[0]!;
            const transKeys = Object.keys(firstTrans);
            // Write headers in B6, C6, ...
            sheet.getRangeByIndexes(5, 1, 1, transKeys.length).values = [transKeys];
            let hasPlaidDataHeaders = false;

            // Write each transaction row individually, resilient to varying column counts
            for (let t = 0; t < result.transactions.length; t++) {
                const trans = result.transactions[t]!;
                const transFieldStrs = transKeys.map((k) => {
                    if (k === "plaid_metadata") {
                        return trans[k] ? "➜" : "";
                    }
                    const val = trans[k];
                    if (val === null) {
                        return "";
                    }
                    // eslint-disable-next-line @typescript-eslint/no-base-to-string
                    return String(val);
                });
                // Write the number of elements in the first column
                sheet.getRangeByIndexes(6 + t, 0, 1, 1).values = [[transFieldStrs.length]];
                // Write the row starting at column B
                sheet.getRangeByIndexes(6 + t, 1, 1, transFieldStrs.length).values = [transFieldStrs];

                await context.sync();

                // Parse Plaid data:
                const plaidDataStr = trans["plaid_metadata"];
                let plaidData = null;

                if (plaidDataStr && typeof plaidDataStr === "string") {
                    plaidData = JSON.parse(plaidDataStr);
                }

                // Output Plaid data:
                if (plaidData) {
                    const plaidKeys = Object.keys(plaidData);
                    if (!hasPlaidDataHeaders) {
                        const plaidKeyHeads = plaidKeys.map((k) => `plaid:${k}`);
                        sheet.getRangeByIndexes(5, 2 + transKeys.length, 1, plaidKeys.length).values = [
                            plaidKeyHeads,
                        ];
                        hasPlaidDataHeaders = true;
                    }

                    const plaidDataFieldStrs = plaidKeys.map((k) => {
                        const val = plaidData[k];
                        if (val === null) {
                            return "";
                        }

                        return String(val);
                    });

                    sheet.getRangeByIndexes(6 + t, 2 + transKeys.length, 1, plaidDataFieldStrs.length).values = [
                        plaidDataFieldStrs,
                    ];
                } else {
                    sheet.getRangeByIndexes(6 + t, 2 + transKeys.length, 1, 1).values = [["NO Plaid data"]];
                }

                await context.sync();
            }
        }
        await context.sync();
    });
}
