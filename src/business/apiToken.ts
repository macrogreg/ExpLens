/// <reference types="office-js" />

import { PromiseCompletionSource } from "src/util/PromiseCompletionSource";
import { isNotNullOrWhitespaceStr } from "src/util/string_util";

// This should be the same as the manifest ID.
const AddInId = "CC923F2C-0638-4F36-9E18-A4910CD71B74";
const TokenSettingName = `${AddInId}.v1ApiToken`;

let apiToken: string | null = null;

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

export function useApiToken() {
    return {
        value: () => {
            if (apiToken !== null) {
                return apiToken;
            }

            const loadedToken = Office.context.document.settings.get(TokenSettingName);
            if (isNotNullOrWhitespaceStr(loadedToken)) {
                apiToken = loadedToken.toString();
                console.log(`LunchMoney API Token loaded from the document (${apiToken.length} chars).`);
            } else {
                console.log("LunchMoney API Token was NOT loaded from the document.");
            }

            return apiToken;
        },

        set: (token: string) => {
            apiToken = token;
            Office.context.document.settings.set(TokenSettingName, apiToken);
        },

        isSet: () => isNotNullOrWhitespaceStr(apiToken),

        store: (): Promise<void> => {
            console.debug("Storing the API Token in the document.");
            Office.context.document.settings.set(TokenSettingName, apiToken);
            return saveDocumentSettings();
        },

        clearStorage: (): Promise<void> => {
            console.debug("Clearing the API Token from the document.");
            Office.context.document.settings.remove(TokenSettingName);
            return saveDocumentSettings();
        },
    };
}
