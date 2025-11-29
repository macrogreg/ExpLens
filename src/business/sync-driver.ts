/// <reference types="office-js" />

import { useSettings } from "src/composables/settings";
import { formatDateUtc } from "src/util/format_util";
import { isNullOrWhitespace } from "src/util/string_util";
import type { TagInfo } from "./tags";
import { downloadTags, SheetNameTags, type TagValuesCollection } from "./tags";
import { downloadCategories, SheetNameCategories } from "./categories";
import { downloadTransactions, SheetNameTransactions } from "./transactions";
import { ensureSheetActive } from "./excel-util";
import type { Ref } from "vue";

export type SyncContext = {
    excel: Excel.RequestContext;
    currentSync: { version: number; utc: Date };
    isReplaceExistingTransactions: boolean;
    progressPercentage: Ref<number>;
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

let isSyncInProgress = false;

export async function downloadData(
    startDate: Date,
    endDate: Date,
    replaceExistingTransactions: boolean,
    syncOperationProgressPercentage: Ref<number>
): Promise<void> {
    if (isSyncInProgress === true) {
        throw new Error("Cannot star data download, because data sync is already in progress.");
    }

    try {
        isSyncInProgress = true;

        console.log(
            `Starting downloadData(startDate=${formatDateUtc(startDate)}, endDate=${formatDateUtc(endDate)},` +
                ` replaceExistingTransactions=${replaceExistingTransactions}).`
        );

        const loadedAppSettings = await useSettings();
        const currentSync = { version: loadedAppSettings.lastCompletedSyncVersion.value + 1, utc: new Date() };

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
                currentSync,
                isReplaceExistingTransactions: replaceExistingTransactions,
                progressPercentage: syncOperationProgressPercentage,
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
            await downloadTransactions(startDate, endDate, syncCtx);
        });

        loadedAppSettings.lastCompletedSyncUtc.value = currentSync.utc;
        loadedAppSettings.lastCompletedSyncVersion.value = currentSync.version;

        console.log(`Completed downloadData(..).`);
    } finally {
        syncOperationProgressPercentage.value = 100;
        isSyncInProgress = false;
    }
}
