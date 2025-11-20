import { errorTypeMessageString } from "src/util/format_util";
import { isNotNullOrWhitespaceStr } from "src/util/string_util";
import { useSettings } from "src/composables/settings";

export async function badResponseToConsole(response: Response, purposeDescription: string | null) {
    let responseText: string;
    let hasResponseText = false;
    try {
        responseText = await response.text();
        hasResponseText = true;
    } catch (err) {
        responseText = `Error getting response text (${errorTypeMessageString(err)})`;
    }

    let responseObject: unknown = { _: "Response text not available for parsing." };
    if (hasResponseText) {
        try {
            responseObject = JSON.parse(responseText);
        } catch (err) {
            responseObject = { errorParsingResponseText: errorTypeMessageString(err) };
        }
    }

    const purpDescr = purposeDescription ? `(${purposeDescription}) ` : "";
    console.error(
        `Request ${purpDescr}failed with status ${response.status} (${response.statusText}).`,
        "Response Text:",
        responseText,
        "Response Object:",
        responseObject
    );
}

export async function authorizedFetch(method: string, api: string, purposeDescription: string): Promise<string> {
    const apiToken = (await useSettings()).apiToken.value;

    if (!isNotNullOrWhitespaceStr(apiToken)) {
        throw new Error(`Cannot '${purposeDescription}', because no API Token is set.`);
    }

    const headers = new Headers();
    headers.append("Authorization", `Bearer ${apiToken}`);
    const requestUrl = `https://dev.lunchmoney.app/v1/${api}`;

    try {
        const response = await fetch(requestUrl, {
            method,
            headers,
            redirect: "follow",
        });

        if (!response.ok) {
            await badResponseToConsole(response, purposeDescription);
            throw new Error(`Bad response (${response.status}) during '${purposeDescription}'.`);
        }

        const responseText = await response.text();
        return responseText;
    } catch (err) {
        console.error(`Fetch error during ${method} /${api} (${purposeDescription}).`, err);
        throw err;
    }
}
