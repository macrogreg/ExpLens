<template>
    <q-page class="row justify-center">
        <div class="col-8 q-pa-md" style="max-width: 800px">
            <div class="q-mb-md">
                <h4 class="q-my-sm">LunchMoney Excel Add-In</h4>
            </div>

            <q-tabs v-model="tab" align="left" class="q-mb-md">
                <q-tab name="syncData" label="Sync Data" />
                <q-tab name="advanced" label="Advanced" />
            </q-tabs>

            <q-separator class="q-mb-md" />

            <sync-data-tab v-if="tab === 'syncData'" />
            <advanced-tab v-if="tab === 'advanced'" />
        </div>
    </q-page>
</template>

<script setup lang="ts">
import { ref } from "vue";
import SyncDataTab from "components/SyncDataTab.vue";
import AdvancedTab from "components/AdvancedTab.vue";

const tab = ref("syncData");
</script>
