/// <reference types="office-js" />

import { errorTypeMessageString } from "src/util/format_util";

function errorMessage(mainMessage: string) {
    return `Cannot initialize Office APIs: ${mainMessage} \nAre you viewing this within the Excel side panel?`;
}

function logIfRequested(isLogRequested: boolean, ...vals: unknown[]) {
    if (isLogRequested) {
        console.error(...vals);
    }
}

export async function useOffice(
    logToConsole: boolean = false
): Promise<{ host: Office.HostType; platform: Office.PlatformType }> {
    if (typeof Office === "undefined") {
        const errMsg = errorMessage("`Office` is not defined in any loaded module.");
        logIfRequested(logToConsole, errMsg);
        throw new Error(errMsg);
    }

    if (Office === undefined || Office === null) {
        const errMsg = errorMessage("`Office` is defined as `null` or `undefined`.");
        logIfRequested(logToConsole, errMsg, "Office:", Office);
        throw new Error(errMsg);
    }

    if (!("onReady" in Office) || typeof Office.onReady !== "function") {
        const errMsg = errorMessage("`Office` does not have an `onReady` property, or it is not a function.");
        logIfRequested(logToConsole, errMsg, "Office:", Office);
        throw new Error(errMsg);
    }

    let officeInfo;
    try {
        officeInfo = await Office.onReady();
    } catch (err) {
        const errMsg = errorMessage(`Error while getting Office ready: '${errorTypeMessageString(err)}'.`);
        logIfRequested(logToConsole, errMsg, "Error:", err);
        throw new Error(errMsg);
    }

    if (officeInfo.host === null && officeInfo.platform === null) {
        const errMsg = errorMessage("API is loaded and ready, but no suitable environment was detected.");
        logIfRequested(logToConsole, errMsg);
        throw new Error(errMsg);
    }

    if (logToConsole) {
        console.info("Office APIs are ready.", "onReadyInfo:", officeInfo);
    }

    return officeInfo;
}
