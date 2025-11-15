<template>
    <q-page class="row justify-center">
        <div class="col-8 q-pa-md" style="max-width: 800px">
            <div class="row items-center no-wrap title-bar">
                <h1 class="col">LunchMoney Sync</h1>
                <div class="col-auto">
                    <img src="/icons/favicon-32x32.png" alt="Logo" style="height: 32px; width: 32px" />
                </div>
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

<style scoped lang="css">
.title-bar {
    background-color: rgb(31, 105, 194);
    padding: 2px 15px 2px 15px;
    color: white;
}
h1 {
    font-family: "Gill Sans", "Gill Sans MT", Calibri, "Trebuchet MS", sans-serif;
    font-weight: 400;
    font-size: 32px;
    line-height: normal;
    margin: 0;
    padding: 5px 0 5px 0;
}
</style>

<script setup lang="ts">
import { ref } from "vue";
import SyncDataTab from "components/SyncDataTab.vue";
import AdvancedTab from "components/AdvancedTab.vue";

const tab = ref("syncData");
</script>
