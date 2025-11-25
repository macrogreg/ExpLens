/// <reference types="office-js" />

import { useSettings } from "src/composables/settings";
import { formatDateUtc } from "src/util/format_util";
import { isNullOrWhitespace } from "src/util/string_util";
import type { TagInfo } from "./tags";
import { downloadTags, SheetNameTags, type TagValuesCollection } from "./tags";
import { downloadCategories, SheetNameCategories } from "./categories";
import { downloadTransactions, SheetNameTransactions } from "./transactions";
import { ensureSheetActive } from "./excel-util";

export type SyncContext = {
    excel: Excel.RequestContext;
    sheets: {
        tags: Excel.Worksheet;
        cats: Excel.Worksheet;
        trans: Excel.Worksheet;
    };
    tags: {
        assignable: TagValuesCollection;
        groupListFormulaLocations: Map<string, string>;
        byId: Map<number, TagInfo>;
    };
};

export async function downloadData(
    startDate: Date,
    endDate: Date,
    replaceExistingTransactions: boolean
): Promise<void> {
    console.log(
        `Starting downloadData(startDate=${formatDateUtc(startDate)}, endDate=${formatDateUtc(endDate)},` +
            ` replaceExistingTransactions=${replaceExistingTransactions}).`
    );

    const loadedAppSettings = await useSettings();

    {
        const apiToken = loadedAppSettings.apiToken.value;
        if (isNullOrWhitespace(apiToken)) {
            console.log("No API token. Cannot proceed with download");
            return;
        }

        console.debug(`downloadData(..): has API token (${apiToken!.length} chars).`);
    }

    await Excel.run(async (context: Excel.RequestContext) => {
        // We need to ensure sheets creation ion the order we want them to appear in the document:
        const transSheet = await ensureSheetActive(SheetNameTransactions, context);
        const tagsSheet = await ensureSheetActive(SheetNameTags, context);
        const catsSheets = await ensureSheetActive(SheetNameCategories, context);

        const syncCtx: SyncContext = {
            excel: context,
            sheets: {
                trans: transSheet,
                tags: tagsSheet,
                cats: catsSheets,
            },
            tags: {
                assignable: new Map<string, Set<string>>(),
                groupListFormulaLocations: new Map<string, string>(),
                byId: new Map<number, TagInfo>(),
            },
        };

        await downloadTags(syncCtx);
        await downloadCategories(context);
        await downloadTransactions(startDate, endDate, context);
    });

    loadedAppSettings.lastCompletedSyncUtc.value = new Date();
    loadedAppSettings.lastCompletedSyncVersion.value = loadedAppSettings.lastCompletedSyncVersion.value + 1;

    console.log(`Completed downloadData(..).`);
}
