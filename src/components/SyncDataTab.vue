<template>
    <div class="text-left q-gutter-md">
        <h5 class="q-mt-none">Sync Data</h5>

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
import { ref } from "vue";

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
const fromDate = ref("");
const toDate = ref("");
const fromDateError = ref("");
const toDateError = ref("");

const tokenRequiredRule = (val: string) =>
    (val && val.trim().length > 0) || "API token must not be empty or whitespace.";

function validateAndDownload() {
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
    // Download handler to be implemented
}
</script>
