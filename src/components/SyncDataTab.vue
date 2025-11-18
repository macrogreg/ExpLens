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

        <div class="q-pa-sm" style="border: 1px lightgray solid; width: fit-content">
            <q-input
                ref="apiTokenTextfield"
                filled
                label="LunchMoney API Token"
                v-model="apiToken"
                :rules="[apiTokenRequiredRule]"
                :dense="true"
                :counter="false"
                maxlength="200"
                style="max-width: 450px; width: 100%; padding: 0 10px 20px 0"
            />

            <q-checkbox
                v-model="hasPersistApiTokenPermissionControl"
                label="Store the API Token in the current documents (Unsecure!)"
                style="font-size: smaller"
            />
        </div>

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

    <q-dialog v-model="showPersistApiTokenDialog" persistent>
        <q-card>
            <q-card-section>
                <div class="text-weight-bold q-mb-sm" style="font-size: larger">
                    Really store the API Token as clear text in this document?
                </div>
                <p class="text-weight-bold">Anybody who can access this document can also access the token.</p>
                <p class="text-justify q-mb-xs">
                    The API Token enables complete access to all your data inside of Lunch Money.<br />
                    We can store the API Token in the current document for your convenience. However, the Token is not
                    encrypted, and anybody with access to this document can theoretically also access the Token.
                </p>
                <p class="text-justify q-mb-xs">
                    If you ever suspect that an unauthorized person accessed your API Token, you must immediately delete it
                    (you can create a new one right away). To do that, go to
                    <span class="text-italic">Settings > Developers</span> in your Lunch Money app.<br />
                    (<a target="_blank" href="https://my.lunchmoney.app/developers">https://my.lunchmoney.app/developers</a
                    >).
                </p>
            </q-card-section>
            <q-card-actions align="right">
                <q-btn flat label="No" color="positive" v-close-popup @click="confirmPersistApiTokenDialog('no')" />
                <q-btn flat label="Yes" color="negative" v-close-popup @click="confirmPersistApiTokenDialog('yes')" />
            </q-card-actions>
        </q-card>
    </q-dialog>
</template>

<style scoped></style>

<script setup lang="ts">
import { ref, onMounted, computed } from "vue";
//import { downloadTransactions } from "../business/downloadTransactions";
import { useApiToken } from "src/business/apiToken";
import { downloadTags } from "src/business/tags";
import { errorTypeMessageString } from "src/util/format_util";
import { downloadCategories } from "src/business/categories";
import { QInput } from "quasar";

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

const apiTokenTextfield = ref<QInput | null>(null);
const apiToken = ref("");

const showPersistApiTokenDialog = ref(false);
const hasPersistApiTokenPermissionData = ref(false);
const hasPersistApiTokenPermissionControl = computed({
    get: () => hasPersistApiTokenPermissionData.value,
    set: (val: boolean) => {
        if (val) {
            showPersistApiTokenDialog.value = true;
        } else {
            hasPersistApiTokenPermissionData.value = false;
        }
    },
});

function confirmPersistApiTokenDialog(choice: "yes" | "no") {
    if (choice === "yes") {
        hasPersistApiTokenPermissionData.value = true;
    }
    showPersistApiTokenDialog.value = false;
}

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
    hasPersistApiTokenPermissionData.value = apiToken.value.length > 0;
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

const apiTokenRequiredRule = (val: string) => (val && val.trim().length > 0) || "API token must not be empty or whitespace.";

async function validateAndDownload() {
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
    if (!(await apiTokenTextfield.value?.validate()) || fromDateError.value || toDateError.value) {
        return;
    }

    // Set API token

    try {
        useApiToken().set(apiToken.value);

        if (hasPersistApiTokenPermissionControl.value) {
            await useApiToken().store();
        } else {
            // If permission is not given, we clear the token, even if it was persisted with permission earlier:
            await useApiToken().clearStorage();
        }
    } catch (err) {
        console.error("Error setting or storing API token.", err);
    }

    //await downloadTransactions(fromDate.value, toDate.value);
    await downloadTags();
    await downloadCategories();
}
</script>
