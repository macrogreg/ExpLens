import { ref, shallowReadonly } from "vue";
import type { ActiveOpsInfo, LogEntry } from "../models/OperationsTracker";
import { OperationsTracker } from "../models/OperationsTracker";
import { captureConsoleToTracker } from "./console-to-tracker-capture";
import { EventLevelKind } from "../models/EventLevelKind";
import { captureWindowErrorsToTracker } from "./window-error-to-tracker-capture";

const FLAG_VALIDATE_LOG_PRUNING = true as const;

export interface StatusLogEntry {
    timestamp: Date;
    message: string;
    getLineString(): string;
}

export const StatusViewTypes = {
    CurrentState: "Current State",
    FullLog: "Full Log",
};

export type StatusViewType = keyof typeof StatusViewTypes;

export const StatusDisplayModes = {
    Always: "Always",
    Never: "Never",
    DuringImportantOperations: "During Important Operations",
};

export type StatusDisplayMode = keyof typeof StatusDisplayModes;

export function rebuildFullLogView(loggedOps: ReadonlyArray<LogEntry>) {
    const view = loggedOps.map((lo) => lo.entry).join("\n");
    return view;
}

function rebuildCurrentStateView(activeOps: ActiveOpsInfo): string {
    let activeOp = activeOps.iterator.next();
    if (activeOp.done) {
        return "";
    }

    let view: string = activeOp.value[1].activeOpsStackEntry;
    activeOp = activeOps.iterator.next();
    while (!activeOp.done) {
        view = activeOp.value[1].activeOpsStackEntry + "\n" + view;
        activeOp = activeOps.iterator.next();
    }

    return view;
}

// Initialization:

// State:

const opTracker = new OperationsTracker();

const statusView = ref<string>("");

const statusViewType = ref<StatusViewType>("CurrentState");
const captureConsole = ref<boolean>(false);
const captureWindowErr = ref<boolean>(false);
const writeToConsole = ref<boolean>(false);

const displayMode = ref<StatusDisplayMode>("DuringImportantOperations");
const isImportantOperationOngoing = ref<boolean>(false);
const isDisplayRequired = ref<boolean>(false);

// Private state:

let _cancelConsoleCaptureFunc: (() => boolean) | null = null;
let _cancelWindowErrCaptureFunc: (() => boolean) | null = null;

// Getters:

const tracker = () => opTracker;

// Actions:

const setStatusViewType = (viewType: StatusViewType): void => {
    if (statusViewType.value === viewType) {
        return;
    }

    const view =
        viewType === "FullLog" ? rebuildFullLogView(opTracker.loggedOps) : rebuildCurrentStateView(opTracker.activeOps);

    statusViewType.value = viewType;
    statusView.value = view;
};

const setCaptureConsole = (capture: boolean): void => {
    //console.debug(`StatusLogState.setCaptureConsole(${capture}): prevVal=${captureConsole.value}`);

    if (captureConsole.value === capture) {
        return;
    }

    if (capture === true) {
        // If there is a stale cancel-capture func, execute it before setting up a new capture:
        if (_cancelConsoleCaptureFunc !== null) {
            _cancelConsoleCaptureFunc();
        }

        const captureHandle = captureConsoleToTracker(opTracker);
        _cancelConsoleCaptureFunc = () => captureHandle.cancel();
    } else {
        if (_cancelConsoleCaptureFunc !== null) {
            _cancelConsoleCaptureFunc();
        }
        _cancelConsoleCaptureFunc = null;
    }

    //console.debug(`StatusLogState.setCaptureConsole: setting captureConsole.value to ${capture}`);
    captureConsole.value = capture;
};

const setCaptureWindowErr = (capture: boolean): void => {
    //console.debug(`StatusLogState.setCaptureWindowErr(${capture}): prevVal=${captureWindowErr.value}`);

    if (captureWindowErr.value === capture) {
        return;
    }

    if (capture === true) {
        // If there is a stale cancel-capture func, execute it before setting up a new capture:
        if (_cancelWindowErrCaptureFunc !== null) {
            _cancelWindowErrCaptureFunc();
        }

        const captureHandle = captureWindowErrorsToTracker(opTracker, { errors: true, unhandledRejection: true });
        _cancelWindowErrCaptureFunc = () => captureHandle.cancel();
    } else {
        if (_cancelWindowErrCaptureFunc !== null) {
            _cancelWindowErrCaptureFunc();
        }
        _cancelWindowErrCaptureFunc = null;
    }

    //console.debug(`StatusLogState.setCaptureWindowErr: setting captureWindowErr.value to ${capture}`);
    captureWindowErr.value = capture;
};

const setWriteToConsole = (write: boolean): void => {
    //console.debug(`StatusLogState.setWriteToConsole(${write}): prevVal=${writeToConsole.value}`);

    if (writeToConsole.value === write) {
        return;
    }

    if (!write) {
        opTracker.observeEvent(EventLevelKind.Inf, "Disabling mirroring tracked log to console.");
    }

    opTracker.config.writeToConsole = write;

    //console.debug(`StatusLogState.setWriteToConsole: setting writeToConsole.value to ${write}`);
    writeToConsole.value = write;

    if (write) {
        opTracker.observeEvent(EventLevelKind.Inf, "Enabled mirroring tracked log to console.");
    }
};

const setDisplayMode = (mode: StatusDisplayMode) => {
    if (displayMode.value === mode) {
        return;
    }

    displayMode.value = mode;
    isDisplayRequired.value =
        displayMode.value === "Always" ||
        (displayMode.value === "DuringImportantOperations" && isImportantOperationOngoing.value === true);
};

const setImportantOperationOngoing = (isOngoing: boolean) => {
    if (isImportantOperationOngoing.value === isOngoing) {
        return;
    }

    isImportantOperationOngoing.value = isOngoing;
    isDisplayRequired.value =
        displayMode.value === "Always" ||
        (displayMode.value === "DuringImportantOperations" && isImportantOperationOngoing.value === true);
};

const notifyLogEntryEmitted = (emit: LogEntry): void => {
    if (statusViewType.value !== "FullLog") {
        return;
    }

    let view = statusView.value;
    view = view + (view.length > 0 ? "\n" : "") + emit.entry;
    statusView.value = view;
};

const notifyLogEntriesDeleted = (removedEntries: LogEntry[], replacementEntries: LogEntry[] | null): void => {
    if (statusViewType.value !== "FullLog") {
        return;
    }

    let view = statusView.value;

    for (let e = 0; e < removedEntries.length; e++) {
        const entry = removedEntries[e]?.entry + "\n";
        const entryLen = entry.length;

        if (FLAG_VALIDATE_LOG_PRUNING) {
            if (!view.startsWith(entry)) {
                // We want to log the error, but we will do it async, after we finished changing the logger data.
                queueMicrotask(() => {
                    opTracker.observeEvent(
                        EventLevelKind.Err,
                        `Status Log State: Error while pruning Log View Cache`,
                        `\n` +
                            ` 🛑  Cannot prune entry #${e} of ${removedEntries.length} removed entries.\n` +
                            `    Entry (len=${entryLen}) "${entry}".\n` +
                            `    But the View Cache does not start with those chars.\n` +
                            `    The first ${entryLen * 2} chars of the View Cache: "${view.slice(0, entryLen * 2)}".\n`
                    );
                });

                // No point trying to prune any more, but we will still add the replacements:
                break;
            }
        }

        view = view.slice(entryLen);
    }

    if (replacementEntries && replacementEntries.length > 0) {
        let replacement = "";
        for (const replace of replacementEntries) {
            replacement += replace.entry + "\n";
        }

        view = replacement + view;
    }

    statusView.value = view;
};

const notifyLogEntriesRevised = (newLog: LogEntry[]): void => {
    if (statusViewType.value !== "FullLog") {
        return;
    }

    const view = rebuildFullLogView(newLog);
    statusView.value = view;
};

const notifyActiveOpsStackUpdated = (activeOps: ActiveOpsInfo): void => {
    if (statusViewType.value !== "CurrentState") {
        return;
    }

    statusView.value = rebuildCurrentStateView(activeOps);
};

const statusLogState = {
    // statusViewType: computed({
    //     get: (): StatusViewType => statusViewType.value,
    //     set: (newVal: StatusViewType) => setStatusViewType(newVal),
    // }),

    // captureConsole: computed({
    //     get: (): boolean => captureConsole.value,
    //     set: (newVal: boolean) => setCaptureConsole(newVal),
    // }),

    // writeToConsole: computed({
    //     get: (): boolean => writeToConsole.value,
    //     set: (newVal: boolean) => setWriteToConsole(newVal),
    // }),

    get tracker() {
        return tracker();
    },

    statusView: shallowReadonly(statusView),

    statusViewType: shallowReadonly(statusViewType),
    captureConsole: shallowReadonly(captureConsole),
    captureWindowErr: shallowReadonly(captureWindowErr),
    writeToConsole: shallowReadonly(writeToConsole),

    displayMode: shallowReadonly(displayMode),
    isImportantOperationOngoing: shallowReadonly(isImportantOperationOngoing),
    isDisplayRequired: shallowReadonly(isDisplayRequired),

    setStatusViewType,
    setCaptureConsole,
    setCaptureWindowErr,
    setWriteToConsole,
    setDisplayMode,
    setImportantOperationOngoing,

    notifyLogEntryEmitted,
    notifyLogEntriesDeleted,
    notifyLogEntriesRevised,
    notifyActiveOpsStackUpdated,
};

opTracker.config.operationsListener = statusLogState;

// Default setting must be applied after everything else is initialized,
// so that the side effects of the respective setters can be executed:
statusLogState.setDisplayMode("Always");
statusLogState.setStatusViewType("FullLog");
statusLogState.setCaptureConsole(true);
statusLogState.setCaptureWindowErr(true);
statusLogState.setWriteToConsole(true);

export function useStatusLog() {
    return statusLogState;
}

export function useOpTracker() {
    return useStatusLog().tracker;
}
