/// <reference types="office-js" />

import { PromiseCompletionSource } from "src/util/PromiseCompletionSource";
import { isNotNullOrWhitespaceStr, isNullOrWhitespace } from "src/util/string_util";
import { useOffice } from "./office-ready";
import { readonly, ref, watch } from "vue";
import { type Ref } from "vue";

// This should be the same as the manifest ID.
const AddInId = "CC923F2C-0638-4F36-9E18-A4910CD71B74";
const ConfigSettingName = `${AddInId}.config`;
const TokenSettingName = `${AddInId}.v1ApiToken`;

// Interface for the exported settings API:
// (see `useSettings()`)
export interface AppSettings {
    appVersion: Readonly<Ref<string>>;
    lastCompletedSyncUtc: Ref<Date | null>;
    lastCompletedSyncVersion: Ref<number>;

    apiToken: Ref<string | null, string | null>;

    storeApiToken: () => Promise<void>;
    clearAllStorage: () => Promise<void>;
    clearTokenStorage: () => Promise<void>;
}

// Interface for storing in the Office document:
interface DocumentConfig {
    appVersion: string;
    lastCompletedSyncUtc: string;
    lastCompletedSyncVersion: number;
}

// Do not await to ensure Office init does not delay module loading:
const settings = initDocumentSettings();

// let isSettingsResolved: boolean = false;
// void settings.then(() => {
//     isSettingsResolved = true;
// });

function saveDocumentSettings(): Promise<void> {
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

    const config: DocumentConfig = (() => {
        const loadedVal = Office.context.document.settings.get(ConfigSettingName) as DocumentConfig;
        if (loadedVal) {
            console.log("DocumentConfig loaded: " + JSON.stringify(loadedVal, null, 4));
            return loadedVal;
        } else {
            console.log("DocumentConfig not found. Will use defaults.");
            return {
                appVersion: "",
                lastCompletedSyncUtc: "",
                lastCompletedSyncVersion: 0,
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

    let lastCompletedSyncUtcDate: Date | null = isNullOrWhitespace(config.lastCompletedSyncUtc)
        ? null
        : new Date(config.lastCompletedSyncUtc);
    if (lastCompletedSyncUtcDate !== null && isNaN(lastCompletedSyncUtcDate.getTime())) {
        console.error(`lastCompletedSyncUtc (='${config.lastCompletedSyncUtc}') cannot be parsed into a valid date.`);
        lastCompletedSyncUtcDate = null;
    }

    const settingsRefs = {
        appVersion: ref<string>(config.appVersion),
        lastCompletedSyncUtc: ref<Date | null>(lastCompletedSyncUtcDate),
        lastCompletedSyncVersion: ref<number>(config.lastCompletedSyncVersion),
        apiToken: ref<string | null>(apiToken),
    };

    watch(settingsRefs.appVersion, async (newVal) => {
        const allSettings = await settings;
        const config: DocumentConfig = {
            appVersion: newVal,
            lastCompletedSyncUtc: allSettings.lastCompletedSyncUtc.value
                ? allSettings.lastCompletedSyncUtc.value.toISOString()
                : "",
            lastCompletedSyncVersion: allSettings.lastCompletedSyncVersion.value,
        };
        Office.context.document.settings.set(ConfigSettingName, config);
        await saveDocumentSettings();
    });

    watch(settingsRefs.lastCompletedSyncUtc, async (newVal) => {
        const allSettings = await settings;
        const config: DocumentConfig = {
            appVersion: allSettings.appVersion.value,
            lastCompletedSyncUtc: newVal === null ? "" : newVal.toISOString(),
            lastCompletedSyncVersion: allSettings.lastCompletedSyncVersion.value,
        };
        Office.context.document.settings.set(ConfigSettingName, config);
        await saveDocumentSettings();
    });

    watch(settingsRefs.lastCompletedSyncVersion, async (newVal) => {
        const allSettings = await settings;
        const config: DocumentConfig = {
            appVersion: allSettings.appVersion.value,
            lastCompletedSyncUtc: allSettings.lastCompletedSyncUtc.value
                ? allSettings.lastCompletedSyncUtc.value.toISOString()
                : "",
            lastCompletedSyncVersion: newVal,
        };
        Office.context.document.settings.set(ConfigSettingName, config);
        await saveDocumentSettings();
    });

    // Do not watch and reactively store the API token. It requires an explicit store invocation.

    return settingsRefs;
}

export async function useSettings(): Promise<AppSettings> {
    const allSettings = await settings;
    return {
        appVersion: readonly(allSettings.appVersion),
        lastCompletedSyncUtc: allSettings.lastCompletedSyncUtc,
        lastCompletedSyncVersion: allSettings.lastCompletedSyncVersion,
        apiToken: allSettings.apiToken,

        storeApiToken: (): Promise<void> => {
            console.log("Storing API token in the document.");
            Office.context.document.settings.set(TokenSettingName, allSettings.apiToken.value);
            return saveDocumentSettings();
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
