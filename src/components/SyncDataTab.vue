<template>
    <div class="text-left q-gutter-md">
        <h5 class="q-ma-md">Sync Data</h5>

        <div class="q-pa-sm" style="border: 1px lightgray solid; font-size: smaller">
            <div v-if="officeApiInitErrorMsg" style="color: red">{{ officeApiInitErrorMsg }}</div>
            <div v-else-if="officeApiEnvInfo">
                Connected to MS Office. Host: '{{ officeApiEnvInfo.host ?? "null" }}'; Platform: '{{
                    officeApiEnvInfo.platform ?? "null"
                }}'.
            </div>
            <div v-else>Office Add-In environment not yet initialized.</div>
        </div>

        <q-input
            filled
            label="LunchMoney API Token"
            v-model="apiToken"
            :error="apiTokenError"
            :rules="[tokenRequiredRule]"
            :dense="true"
            :counter="false"
            maxlength="200"
            style="max-width: 450px; width: 100%; padding-right: 10px"
        />
        <!-- Error message handled by Quasar input rule -->

        <div class="date-inputs" style="display: flex; gap: 16px">
            <div style="max-width: 220px; width: 100%">
                <q-input
                    filled
                    label="FROM"
                    v-model="fromDate"
                    type="date"
                    :error="!!fromDateError"
                    :rules="[fromDateAgeRule]"
                />
                <div v-if="fromDateError" class="text-negative q-mt-xs">{{ fromDateError }}</div>
            </div>
            <div style="max-width: 220px; width: 100%">
                <q-input
                    filled
                    label="TO"
                    v-model="toDate"
                    type="date"
                    :error="!!toDateError"
                    :rules="[toDateOrderRule, toDateAgeRule]"
                />
                <div v-if="toDateError" class="text-negative q-mt-xs">{{ toDateError }}</div>
            </div>
        </div>

        <q-btn label="Download" color="primary" @click="validateAndDownload" class="q-mt-md" />
    </div>
</template>

<script setup lang="ts">
import { ref, onMounted } from "vue";
//import { downloadTransactions } from "../business/downloadTransactions";
import { useApiToken } from "src/business/apiToken";
import { downloadTags } from "src/business/tags";
import { errorTypeMessageString } from "src/util/format_util";

const officeApiInitErrorMsg = ref("");
const officeApiEnvInfo = ref<null | { host: Office.HostType; platform: Office.PlatformType }>(null);

const now = new Date();
const twoYearsAgo = new Date(now.getTime() - 730 * 24 * 60 * 60 * 1000);

const fromDateAgeRule = (val: string) => {
    if (!val) return true;
    const from = new Date(val);
    return from >= twoYearsAgo || "FROM date must be within 2 years.";
};

const toDateAgeRule = (val: string) => {
    if (!val) return true;
    const to = new Date(val);
    return to >= twoYearsAgo || "TO date must be within 2 years.";
};

const toDateOrderRule = (val: string) => {
    if (!val || !fromDate.value) return true;
    const from = new Date(fromDate.value);
    const to = new Date(val);
    return to >= from || "TO date cannot be before FROM date.";
};

const apiToken = ref("");
const apiTokenError = ref(false);

onMounted(async () => {
    console.debug("LunchMoney Excel-AddIn: SyncData Tab mounted. Getting Office API ready...");

    if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
        try {
            officeApiEnvInfo.value = await Office.onReady();
        } catch (err) {
            console.error(
                "LunchMoney Excel-AddIn: Failed initializing Office AddIn environment." +
                    " Are you viewing this within the Excel side panel?",
                err
            );
            officeApiInitErrorMsg.value =
                "Failed initializing Office AddIn environment." +
                " Are you viewing this in an Excel Add-In Side Panel? You must!" +
                " Diagnostic details: " +
                errorTypeMessageString(err);
            return;
        }
    } else {
        console.error(
            "LunchMoney Excel-AddIn: Cannot initialize Office APIs: `Office.onReady(..)` is not available." +
                " Are you viewing this within the Excel side panel?"
        );
        officeApiInitErrorMsg.value =
            "Cannot initialize Office APIs." +
            " Are you viewing this in an Excel Add-In Side Panel? You must!" +
            " Diagnostic details: " +
            "`Office.onReady(..)` is not available.";
        return;
    }

    if (officeApiEnvInfo.value.host === null && officeApiEnvInfo.value.platform === null) {
        console.error(
            "LunchMoney Excel-AddIn: Office AddIn initialization completed, but no suitable environment was detected." +
                " Are you viewing this within the Excel side panel?"
        );

        officeApiInitErrorMsg.value =
            "Office AddIn initialization completed, but no suitable environment was detected." +
            " Are you viewing this in an Excel Add-In Side Panel? You must!";
        officeApiEnvInfo.value = null;
        return;
    }

    console.log("LunchMoney Excel-AddIn: Office API is ready.", officeApiEnvInfo.value);
    officeApiInitErrorMsg.value = "";

    apiToken.value = useApiToken().value() ?? "";
});

// Helper to format local date as YYYY-MM-DD
function formatLocalDate(date: Date): string {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

const today = new Date();
let defaultFromDate: string;
if (today.getDate() >= 1 && today.getDate() <= 19) {
    // 1st of previous month
    const prevMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    defaultFromDate = formatLocalDate(prevMonth);
} else {
    // 1st of current month
    const firstOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
    defaultFromDate = formatLocalDate(firstOfMonth);
}
const defaultToDate = formatLocalDate(today);

const fromDate = ref(defaultFromDate);
const toDate = ref(defaultToDate);
const fromDateError = ref("");
const toDateError = ref("");

const tokenRequiredRule = (val: string) => (val && val.trim().length > 0) || "API token must not be empty or whitespace.";

async function validateAndDownload() {
    apiTokenError.value = !(apiToken.value && apiToken.value.trim().length > 0);

    // Date validation
    fromDateError.value = "";
    toDateError.value = "";

    const from = fromDate.value ? new Date(fromDate.value) : null;
    const to = toDate.value ? new Date(toDate.value) : null;

    if (!fromDate.value) {
        fromDateError.value = "Please select a FROM date.";
    } else if (from && from < twoYearsAgo) {
        fromDateError.value = "FROM date must be within 2 years.";
    }

    if (!toDate.value) {
        toDateError.value = "Please select a TO date.";
    } else if (to && to < twoYearsAgo) {
        toDateError.value = "TO date must be within 2 years.";
    }

    if (from && to && to < from) {
        toDateError.value = "TO date cannot be before FROM date.";
    }

    // Highlight errors, do not proceed if any
    if (apiTokenError.value || fromDateError.value || toDateError.value) {
        return;
    }

    // Set API token

    try {
        useApiToken().set(apiToken.value);
        await useApiToken().store();
    } catch (err) {
        console.error("Error setting or storing API token.", err);
    }

    //await downloadTransactions(fromDate.value, toDate.value);
    await downloadTags();
}
</script>
