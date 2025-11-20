/// <reference types="office-js" />

import { PromiseCompletionSource } from "src/util/PromiseCompletionSource";
import { isNotNullOrWhitespaceStr, isNullOrWhitespace } from "src/util/string_util";
import { useOffice } from "./office-ready";
import { readonly, ref, watch } from "vue";

// This should be the same as the manifest ID.
const AddInId = "CC923F2C-0638-4F36-9E18-A4910CD71B74";
const ConfigSettingName = `${AddInId}.config`;
const TokenSettingName = `${AddInId}.v1ApiToken`;

interface DocumentConfig {
    appVersion: string;
    lastCompletedSync: string;
}

async function saveDocumentSettings() {
    const completion = new PromiseCompletionSource<void>();
    Office.context.document.settings.saveAsync((result: Office.AsyncResult<void>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            completion.tryResolve();
        } else {
            completion.tryReject(result.error);
        }
    });

    return completion.promise();
}

async function initDocumentSettings() {
    // Make sure Office APIs are available:
    await useOffice();

    const config = (() => {
        const loadedVal = Office.context.document.settings.get(ConfigSettingName) as DocumentConfig;
        if (loadedVal) {
            console.log("DocumentConfig loaded: " + JSON.stringify(loadedVal, null, 4));
            return loadedVal;
        } else {
            console.log("DocumentConfig not found. Will use defaults.");
            return {
                appVersion: "",
                lastCompletedSync: "",
            };
        }
    })();

    const apiToken = (() => {
        const loadedVal = Office.context.document.settings.get(TokenSettingName);
        if (isNotNullOrWhitespaceStr(loadedVal)) {
            console.log(`LunchMoney API Token loaded from the document (${loadedVal.length} chars).`);
            return loadedVal.toString();
        } else {
            console.log("LunchMoney API Token was NOT loaded from the document.");
            return null;
        }
    })();

    let lastCompletedSyncDate: Date | null = isNullOrWhitespace(config.lastCompletedSync)
        ? null
        : new Date(config.lastCompletedSync);
    if (lastCompletedSyncDate !== null && isNaN(lastCompletedSyncDate.getTime())) {
        console.error(`lastCompletedSync (='${config.lastCompletedSync}') cannot be parsed into a valid date.`);
        lastCompletedSyncDate = null;
    }

    const settingsRefs = {
        appVersion: ref<string>(config.appVersion),
        lastCompletedSync: ref<Date | null>(lastCompletedSyncDate),
        apiToken: ref<string | null>(apiToken),
    };

    watch(settingsRefs.appVersion, async (newVal) => {
        const allSettings = await settings;
        const config: DocumentConfig = {
            appVersion: newVal,
            lastCompletedSync: allSettings.lastCompletedSync.value ? allSettings.lastCompletedSync.value.toISOString() : "",
        };
        Office.context.document.settings.set(ConfigSettingName, config);
        await saveDocumentSettings();
    });

    watch(settingsRefs.lastCompletedSync, async (newVal) => {
        const allSettings = await settings;
        const config: DocumentConfig = {
            appVersion: allSettings.appVersion.value,
            lastCompletedSync: newVal === null ? "" : newVal.toISOString(),
        };
        Office.context.document.settings.set(ConfigSettingName, config);
        await saveDocumentSettings();
    });

    // Do not watch and reactively store the API token. It required an explicit store invocation.

    return settingsRefs;
}

// Do not await to ensure Office init does not delay module loading:
const settings = initDocumentSettings();

export async function useSettings() {
    return {
        appVersion: readonly((await settings).appVersion),
        lastCompletedSync: (await settings).lastCompletedSync,
        apiToken: (await settings).apiToken,

        storeApiToken: async () => {
            const allSettings = await settings;
            console.log("Storing API token in the document.");
            Office.context.document.settings.set(TokenSettingName, allSettings.apiToken.value);
            await saveDocumentSettings();
        },

        clearAllStorage: (): Promise<void> => {
            console.debug("Clearing all settings from the document store.");
            Office.context.document.settings.remove(ConfigSettingName);
            Office.context.document.settings.remove(TokenSettingName);
            return saveDocumentSettings();
        },

        clearTokenStorage: (): Promise<void> => {
            console.debug("Clearing API Token all settings from the document store.");
            Office.context.document.settings.remove(ConfigSettingName);
            Office.context.document.settings.remove(TokenSettingName);
            return saveDocumentSettings();
        },
    };
}
