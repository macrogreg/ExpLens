<template>
    <div class="text-left q-gutter-md">
        <h5 class="q-ma-md">Sync Data</h5>

        <div class="q-pa-sm" style="border: 1px lightgray solid; font-size: smaller">
            <div v-if="officeApiInitErrorMsg" style="color: red">{{ officeApiInitErrorMsg }}</div>
            <div v-else-if="officeApiEnvInfo">
                <div>
                    Connected to MS Office. Host: '{{ officeApiEnvInfo.host ?? "null" }}'; Platform: '{{
                        officeApiEnvInfo.platform ?? "null"
                    }}'.
                </div>
                <div v-if="appSettings">
                    Last sync:
                    {{
                        appSettings.lastCompletedSyncUtc.value
                            ? formatDateTimeLocalLong(appSettings.lastCompletedSyncUtc.value)
                            : "never"
                    }}
                    (#{{ appSettings.lastCompletedSyncVersion.value }}).
                </div>
            </div>
            <div v-else>Office Add-In environment not yet initialized.</div>
        </div>

        <div class="q-pa-sm" style="border: 1px lightgray solid; width: fit-content">
            <q-expansion-item
                id="api-token-expansion"
                v-model="isApiTokenAreaExpanded"
                label="Lunch Money Access Token"
                :caption="isApiTokenAreaExpanded ? '\u00A0' : apiTokenExpansionCaption"
                dense
                dense-toggle
            >
                <q-input
                    ref="apiTokenTextfield"
                    filled
                    label="API Token"
                    v-model="apiTokenModel"
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

                <div style="padding: 0; margin: 0px; text-align: right">
                    <q-btn
                        :label="'\u00A0Apply\u00A0'"
                        :loading="isDataOperationInProgress"
                        :percentage="syncOperationProgressPercentage"
                        color="primary"
                        @click="checkAndApplyApiToken"
                        class="q-ma-sm"
                        dense
                        no-caps
                    />
                </div>
                <!-- prettier-ignore -->
                <pre
                    class="q-ma-xs"
                    style="font-family: monospace; font-size: 10px; line-height: 1;"
                >{{ apiTokenValidateResultMsg }}</pre>
            </q-expansion-item>
        </div>

        <div class="date-inputs" style="display: flex; gap: 16px">
            <div style="max-width: 220px; width: 100%">
                <q-input
                    filled
                    label="FROM"
                    v-model="syncStartDate"
                    type="date"
                    :error="!!syncStartDateError"
                    :rules="[syncStartDateAgeRule]"
                />
                <div v-if="syncStartDateError" class="text-negative q-mt-xs">{{ syncStartDateError }}</div>
            </div>
            <div style="max-width: 220px; width: 100%">
                <q-input
                    filled
                    label="TO"
                    v-model="syncEndDate"
                    type="date"
                    :error="!!syncEndDateError"
                    :rules="[toDateOrderRule, toDateAgeRule]"
                />
                <div v-if="syncEndDateError" class="text-negative q-mt-xs">{{ syncEndDateError }}</div>
            </div>
        </div>

        <q-btn
            label="Download"
            :loading="isDataOperationInProgress"
            :percentage="syncOperationProgressPercentage"
            color="primary"
            @click="validateAndDownload"
            class="q-mt-md"
        />
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
                    If you ever suspect that an unauthorized person accessed your API Token, you must immediately delete
                    it (you can create a new one right away). To do that, go to
                    <span class="text-italic">Settings > Developers</span> in your Lunch Money app.<br />
                    (<a target="_blank" href="https://my.lunchmoney.app/developers"
                        >https://my.lunchmoney.app/developers</a
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

<style>
#api-token-expansion .q-focus-helper {
    visibility: hidden;
}
</style>

<script setup lang="ts">
import { ref, onMounted, computed, watch } from "vue";
import { formatDateLocal, formatDateTimeLocalLong, formatValue } from "src/util/format_util";
import { QInput } from "quasar";
import { useOffice } from "src/composables/office-ready";
import { type AppSettings, useSettings } from "src/composables/settings";
import { downloadData } from "src/business/sync-driver";
import { useOpTracker } from "src/status-tracker/composables/status-log";
import { validateApiToken } from "src/business/validate-api-token";

const officeApiInitErrorMsg = ref("");
const officeApiEnvInfo = ref<null | { host: Office.HostType; platform: Office.PlatformType }>(null);

let appSettings: AppSettings;

const now = new Date();
const twoYearsAgo = new Date(now.getTime() - 730 * 24 * 60 * 60 * 1000);

const syncStartDateAgeRule = (val: string) => {
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
    if (!val || !syncStartDate.value) return true;
    const from = new Date(syncStartDate.value);
    const to = new Date(val);
    return to >= from || "TO date cannot be before FROM date.";
};

const apiTokenTextfield = ref<QInput | null>(null);
const apiTokenModel = ref("");

const isApiTokenAreaExpanded = ref(true);
const apiTokenValidateResultMsg = ref<string>("API token not verified.");
const isApiTokenValid = ref<boolean | undefined>(undefined);
const isApiTokenValidationInProgress = ref<boolean>(false);

const apiTokenExpansionCaption = computed(() => {
    if (isApiTokenValid.value === undefined) {
        return "Token not verified." + (isApiTokenAreaExpanded.value ? "" : " Expand to configure.");
    }

    return isApiTokenValid.value
        ? "Token is valid"
        : "Token not valid." + (isApiTokenAreaExpanded.value ? "" : " Expand to configure.");
});

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

const apiTokenRequiredRule = (val: string) =>
    (val && val.trim().length > 0) || "API token must not be empty or whitespace.";

async function checkApiToken(token: string) {
    isApiTokenValidationInProgress.value = true;
    try {
        const { isValid, info } = await validateApiToken(token);
        if (!isValid) {
            apiTokenValidateResultMsg.value = info;
        } else {
            apiTokenValidateResultMsg.value =
                "Token is valid." +
                `\n Token moniker:     ${info.api_key_label}` +
                `\n Grants access to:  ${info.budget_name}` +
                `\n User:              ${info.user_name} (${info.user_email})`;
        }
        isApiTokenValid.value = isValid;
        return isValid;
    } finally {
        isApiTokenValidationInProgress.value = false;
    }
}

async function checkAndApplyApiToken(): Promise<string | undefined> {
    if (!(await apiTokenTextfield.value?.validate())) {
        return;
    }

    const token = apiTokenModel.value;
    const isValid = await checkApiToken(token);

    if (!isValid) {
        return undefined;
    }

    try {
        const appSettings = await useSettings();
        const tokenAppSetting = appSettings.apiToken;
        tokenAppSetting.value = token;

        if (hasPersistApiTokenPermissionControl.value) {
            await appSettings.storeApiToken();
        } else {
            // If permission is not given, we clear the token, even if it was persisted with permission earlier:
            // (token will remain in settings, but will not be stored)
            await appSettings.clearTokenStorage();
        }
    } catch (err) {
        // We failed working with storage, but we already validated the token.
        console.error("Error setting or storing API token in App Settings.", err);
    }

    return token;
}

function confirmPersistApiTokenDialog(choice: "yes" | "no") {
    if (choice === "yes") {
        hasPersistApiTokenPermissionData.value = true;
    }
    showPersistApiTokenDialog.value = false;
}

onMounted(async () => {
    const op = useOpTracker().startOperation("ExpLens Excel-AddIn: SyncData Tab mounted. Getting Office API ready...");
    try {
        try {
            officeApiEnvInfo.value = await useOffice(true);
        } catch (err) {
            if (err instanceof Error) {
                officeApiInitErrorMsg.value = err.message;
            } else {
                officeApiInitErrorMsg.value = "Unexpected error while getting office APIs ready: " + formatValue(err);
            }
            op.setFailure("Error getting office APIs ready. AddIn will not work!", {
                message: officeApiInitErrorMsg.value,
                error: err,
            });
            console.error("Error getting office APIs ready. AddIn will not work!", {
                message: officeApiInitErrorMsg.value,
                error: err,
            });
            return;
        }

        officeApiInitErrorMsg.value = "";

        appSettings = await useSettings();

        const apiTokenSetting = appSettings.apiToken;
        apiTokenModel.value = apiTokenSetting.value ?? "";
        hasPersistApiTokenPermissionData.value = apiTokenModel.value.length > 0;

        const isLoadedTokenValid = await checkApiToken(apiTokenModel.value);
        isApiTokenAreaExpanded.value = !isLoadedTokenValid;

        // If token in app settings changes, update the us immediately:
        // (other direction only via apply or sync button)
        watch(apiTokenSetting, (newVal) => {
            apiTokenModel.value = newVal ?? "";
        });

        op.setSuccess();
    } catch (err) {
        op.setFailureAndRethrow(err);
    }
});

function getDefaultSyncStartDate(today: Date) {
    if (today.getDate() >= 1 && today.getDate() <= 19) {
        // 1st of previous month
        const prevMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
        return formatDateLocal(prevMonth);
    } else {
        // 1st of current month
        const firstOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
        return formatDateLocal(firstOfMonth);
    }
}

const syncOperationProgressPercentage = ref<number>(-1);
const isDataOperationInProgress = computed<boolean>(() => {
    const isDataSync = 0 <= syncOperationProgressPercentage.value && syncOperationProgressPercentage.value < 100;
    const isTokenValidation = isApiTokenValidationInProgress.value === true;
    return isDataSync || isTokenValidation;
});

const syncStartDate = ref(getDefaultSyncStartDate(now));
const syncEndDate = ref(formatDateLocal(now));
const syncStartDateError = ref("");
const syncEndDateError = ref("");

async function validateAndDownload() {
    // Date validation
    syncStartDateError.value = "";
    syncEndDateError.value = "";

    const startDate = syncStartDate.value ? new Date(syncStartDate.value) : null;
    const endDate = syncEndDate.value ? new Date(syncEndDate.value) : null;

    if (!syncStartDate.value) {
        syncStartDateError.value = "Please select a FROM date.";
    } else if (startDate && startDate < twoYearsAgo) {
        syncStartDateError.value = "FROM date must be within 2 years.";
    }

    if (!syncEndDate.value) {
        syncEndDateError.value = "Please select a TO date.";
    } else if (endDate && endDate < twoYearsAgo) {
        syncEndDateError.value = "TO date must be within 2 years.";
    }

    if (startDate && endDate && endDate < startDate) {
        syncEndDateError.value = "TO date cannot be before FROM date.";
    }

    // Highlight errors, do not proceed if any
    if (syncStartDateError.value || syncEndDateError.value) {
        return;
    }

    if (startDate === null || endDate === null) {
        return;
    }

    if ((await checkAndApplyApiToken()) === undefined) {
        isApiTokenAreaExpanded.value = true;
        return;
    }

    try {
        syncOperationProgressPercentage.value = 0;
        await downloadData(startDate, endDate, true, syncOperationProgressPercentage);
    } finally {
        syncOperationProgressPercentage.value = -1;
    }
}
</script>
