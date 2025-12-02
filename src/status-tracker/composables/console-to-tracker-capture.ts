//

import { type ConsoleFuncName, createVirtualConsole, redirectConsole } from "../sysutil/ConsoleRedirect";
import { EventLevelKind, type OperationsTracker } from "../models/OperationsTracker";
import { errorTypeMessageString, errorTypeString, formatValueSimple } from "src/util/format_util";
import { DelayPromise } from "src/util/DelayPromise";

export type ConsoleCaptureHandle = {
    isCancelled: () => boolean;
    /** The cancel() function may be called several times, but the cancellation can obviously occur only
     * once. It returns `true` on the first call the the capturing is actually cancelled, and `false` otherwise. */
    cancel: () => boolean;
    invokeOriginalFunc: (funcName: ConsoleFuncName, ...args: unknown[]) => boolean;
};

const CapturedCallMessage = "console" as const;
const BeginCaptureMessage = " *-*-* 🧲 Console Output Capture BEGIN *-*-*" as const;
const EndCaptureMessage = " *-*-* 🧲 Console Output Capture END *-*-*" as const;

export function captureConsoleToTracker(tracker: OperationsTracker): ConsoleCaptureHandle {
    //

    // Capture the console before the redirection, so that the tracker can write to it,
    // and it's output does not get redirected circularly:
    const virtualConsole = createVirtualConsole();
    const prevVirtualConsole = tracker.config.virtualConsole;
    tracker.config.virtualConsole = virtualConsole;

    const errorEventHandler = (errorEvent: Event | ErrorEvent) => {
        const evT = "type" in errorEvent ? errorEvent.type : undefined;
        const msg = "message" in errorEvent ? errorEvent.message : undefined;
        const errT = "error" in errorEvent ? errorTypeString(errorEvent.error) : undefined;
        const errM = "error" in errorEvent ? errorTypeMessageString(errorEvent.error) : undefined;
        const titleInfo = errT ? ` (${errT})` : evT ? ` (${evT})` : "";

        if (msg === ResizeObserverLoopCommonError) {
            handleResizeObserverLoopCommonError(titleInfo, msg ?? errM ?? evT, errorEvent, tracker);
            return;
        }

        tracker.observeEvent(
            EventLevelKind.Err | EventLevelKind.ConsoleCapture,
            `An unhandled error occurred${titleInfo}.`,
            msg ?? errM ?? evT,
            errorEvent
        );
    };

    const asyncErrorEventHandler = (errorEvent: Event | PromiseRejectionEvent) => {
        const evT = "type" in errorEvent ? errorEvent.type : undefined;
        const hasReas = "reason" in errorEvent;
        const reas = hasReas ? errorEvent.reason : undefined;

        const reasT = typeof reas;
        const isErr = reas && reasT === "object" && reas instanceof Error;

        const titleInfo = isErr
            ? ` (${errorTypeString(reas)})`
            : reasT !== "object" && reasT !== "function"
              ? ` (${String(reas)})`
              : "";

        tracker.observeEvent(
            EventLevelKind.Err | EventLevelKind.ConsoleCapture,
            `An unhandled promise rejection occurred${titleInfo}.`,
            hasReas ? formatValueSimple(reas) : evT,
            errorEvent
        );
    };

    window.addEventListener("error", errorEventHandler, true);
    window.addEventListener("unhandledrejection", asyncErrorEventHandler);

    // Redirect console calls so that they form data to the tracker, but still write to the console
    const underlyingRedirHndl = redirectConsole(
        (consFN, ...args) => {
            const eventLevel = mapConsoleFuncToEventKind(consFN);
            const message = `${CapturedCallMessage}.${consFN}`;

            let info;
            if (args.length === 0) {
                info = "-no-args";
            } else {
                const argStr = formatValueSimple(args[0]);
                info = `arg1: ${argStr}; arg count: ${args.length}`;
            }

            tracker.observeEvent(eventLevel, message, info, ...args);
        },
        { invokeOriginals: true }
    );

    // Log that capturing started. Capturing is already active, so this will appear in both, console and tracker.
    console.warn(BeginCaptureMessage);

    // Create a handle that can be used to cancel the capture:
    const captureHandle = {
        isCancelled: () => underlyingRedirHndl.isCancelled(),

        cancel: (): boolean => {
            if (underlyingRedirHndl.isCancelled()) {
                return false;
            }

            // Log that capturing is ending. Capturing is still active, so this will appear in both, console and tracker.
            console.warn(EndCaptureMessage);

            // Stop redirection:
            const hasCanceled = underlyingRedirHndl.cancel();

            // Restore tracker's console to whatever it was before:
            if (hasCanceled && tracker.config.virtualConsole === virtualConsole) {
                window.removeEventListener("unhandledrejection", asyncErrorEventHandler);
                window.removeEventListener("error", errorEventHandler, true);

                tracker.config.virtualConsole = prevVirtualConsole;
            } else {
                const msg =
                    "ConsoleCaptureHandle.cancel(): Original `virtualConsole` of the tracker was not restored" +
                    " because the current virtualConsole is not the one installed by this ConsoleCaptureHandle." +
                    " Was the console redirected multiple times?";
                tracker.observeEvent(EventLevelKind.Wrn, msg);
                virtualConsole.warn(msg);
            }

            return hasCanceled;
        },

        invokeOriginalFunc: (funcName: ConsoleFuncName, ...args: unknown[]) => {
            return underlyingRedirHndl.invokeOriginalFunc(funcName, ...args);
        },
    };

    return captureHandle;
}

const mapConsoleFuncToEventKind = (fn: ConsoleFuncName): EventLevelKind => {
    switch (fn) {
        case "error":
            return EventLevelKind.Err | EventLevelKind.ConsoleCapture;
        case "warn":
            return EventLevelKind.Wrn | EventLevelKind.ConsoleCapture;
        case "log":
            return EventLevelKind.Suc | EventLevelKind.ConsoleCapture;
        default:
            return EventLevelKind.Inf | EventLevelKind.ConsoleCapture;
    }
};

const ResizeObserverLoopCommonError = "ResizeObserver loop completed with undelivered notifications.";
const AccumulatorPeriodMsec = 5000;

let accumulatorResizeObserverLoopError: null | ReturnType<typeof createEventAccumulator> = null;

function handleResizeObserverLoopCommonError(
    titleInfo: string,
    msg: string | undefined,
    errorEvent: Event | ErrorEvent,
    tracker: OperationsTracker
) {
    if (accumulatorResizeObserverLoopError === null) {
        accumulatorResizeObserverLoopError = createEventAccumulator(AccumulatorPeriodMsec, tracker, () => {
            accumulatorResizeObserverLoopError = null;
        });
    }

    accumulatorResizeObserverLoopError.observeEvent(
        EventLevelKind.Err | EventLevelKind.ConsoleCapture,
        `An error occurred in the browser window and was not handled${titleInfo}.`,
        msg,
        errorEvent
    );
}

function createEventAccumulator(lifetimeMsec: number, tracker: OperationsTracker, finishingFn: () => void) {
    let countEvents = 0;

    let lastEvent:
        | undefined
        | {
              kind: EventLevelKind;
              name: string | string[];
              eventInfo: unknown;
              moreInfo: unknown[];
          } = undefined;

    DelayPromise.Run(lifetimeMsec)
        .finally(() => {
            finishingFn();
            if (countEvents === 0) {
                return;
            }

            const lastEv = lastEvent!;
            const n =
                `Multiple events (${countEvents}) of the same kind occurred and were throttled over a period of ` +
                ` ${lifetimeMsec} msec. Last event is shown.`;
            tracker.observeEvent(lastEv.kind, n, lastEv.eventInfo, ...lastEv.moreInfo);
        })
        .catch(() => {});

    return {
        observeEvent: (
            kind: EventLevelKind,
            name: string | string[],
            eventInfo: unknown = undefined,
            ...moreInfo: unknown[]
        ) => {
            if (countEvents === 0) {
                countEvents = 1;
                lastEvent = undefined;
                tracker.observeEvent(kind, name, eventInfo, ...moreInfo);
                return;
            }
            countEvents++;
            lastEvent = { kind, name, eventInfo, moreInfo };
        },
    };
}
