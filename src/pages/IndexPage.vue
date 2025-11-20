<template>
    <q-page class="row justify-center">
        <div class="main-content q-pa-md">
            <div class="title-bar">
                <h1 class="title-text">LunchMoney Sync</h1>
                <img src="/icons/favicon-32x32.png" alt="Logo" class="title-logo" />
            </div>

            <q-tabs v-model="tab" align="left" class="q-mb-md">
                <q-tab name="syncData" label="Sync Data" />
                <q-tab name="advanced" label="Advanced" />
            </q-tabs>

            <q-separator class="q-mb-lg" />

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
    display: flex;
    align-items: center;
}
.title-text {
    font-family: "Gill Sans", "Gill Sans MT", Calibri, "Trebuchet MS", sans-serif;
    font-weight: 400;
    font-size: 32px;
    line-height: normal;
    margin: 0;
    padding: 5px 0 5px 0;
    flex: 1;
}
.title-logo {
    height: 32px;
    width: 32px;
    margin-left: 12px;
}
.main-content {
    width: 100%;
    max-width: 800px;
    margin: 0 auto;
    box-sizing: border-box;
}
</style>

<script setup lang="ts">
import { ref, onMounted } from "vue";

console.debug("LunchMoney Excel-AddIn: Loading IndexPage...");

onMounted(async () => {});
import SyncDataTab from "components/SyncDataTab.vue";
import AdvancedTab from "components/AdvancedTab.vue";

const tab = ref("syncData");
</script>
