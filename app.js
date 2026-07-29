/******************************************************************
 * Workbook Automation Tool
 * app.js V2
 * Part 1A
 * Globals • DOM Cache • Bootstrap • Initialization
 ******************************************************************/
"use strict";

/* ============================================================
   APPLICATION STATE
============================================================ */

const AppState = {

    workbook: null,
    workbookName: "",
    workbookType: "",
    selectedSheet: "",

    selectedSheets: [],

    worksheetSelectionOrder: [],

    monitorId: "",
    monitorTimer: null,
    temporaryWorkbook: null,

    prompt: {
        offer: "",
        context: "",
        sheet: "",
        final: ""
    },

    dirty: false,

    loading: false,

    workflowStep: 1,

    maskingSummary: [],

    promptSafe: false

};

/* ============================================================
   DOM CACHE
============================================================ */

const UI = {

    /* Header */

    projectRef:
        document.getElementById("projectRef"),

    taskNumber:
        document.getElementById("taskNumber"),

    fetchButton:
        document.getElementById("fetchWorkbookBtn"),

    resetWorkspaceButton:
        document.getElementById("resetWorkspaceBtn"),

    manualDownloadPanel:
        document.getElementById("manualDownloadPanel"),

    manualDownloadMessage:
        document.getElementById("manualDownloadMessage"),

    expectedWorkbookName:
        document.getElementById("expectedWorkbookName"),

    openTaskButton:
        document.getElementById("openTaskBtn"),

    manualWorkbookInput:
        document.getElementById("manualWorkbookInput"),

    downloadMonitorStatus:
        document.getElementById("downloadMonitorStatus"),

    cancelMonitorButton:
        document.getElementById("cancelMonitorBtn"),

    manualUploadLabel:
        document.getElementById("manualUploadLabel"),

    /* KPI */

    workbookName:
        document.getElementById("workbookName"),

    workbookType:
        document.getElementById("workbookType"),

    workbookSource:
        document.getElementById("workbookSource"),

    sheetCount:
        document.getElementById("sheetCount"),

    piiStatus:
        document.getElementById("piiStatus"),

    aiStatus:
        document.getElementById("aiStatus"),

    /* Prompt Builder */

    sheetSelector:
        document.getElementById("sheetSelector"),

    worksheetDropdownButton:
        document.getElementById("worksheetDropdownButton"),

    worksheetCheckboxList:
        document.getElementById("worksheetCheckboxList"),

    selectAllWorksheetsButton:
        document.getElementById("selectAllWorksheetsBtn"),

    clearAllWorksheetsButton:
        document.getElementById("clearAllWorksheetsBtn"),

    previewSheetButton:
        document.getElementById("previewSheetBtn"),

    worksheetSelectionMessage:
        document.getElementById("worksheetSelectionMessage"),

    offer:
        document.getElementById("offerDescription"),

    context:
        document.getElementById("workbookContext"),

    workbookContentList:
        document.getElementById("workbookContentList"),

    sheetPrompt:
        document.getElementById("sheetPrompt"),

    finalPrompt:
        document.getElementById("finalPrompt"),

    /* Preview */

    previewButton:
        document.getElementById("previewPromptBtn"),

    previewArea:
        document.getElementById("previewPrompt"),

    previewModal:
        document.getElementById("promptPreviewModal"),

    sheetPreviewModal:
        document.getElementById("sheetPreviewModal"),

    sheetPreviewGrid:
        document.getElementById("sheetPreviewGrid"),

    sheetPreviewTabs:
        document.getElementById("sheetPreviewTabs"),

    sheetPreviewWorkbookName:
        document.getElementById("sheetPreviewWorkbookName"),

    sheetPreviewNotice:
        document.getElementById("sheetPreviewNotice"),

    copyButton:
        document.getElementById("copyPromptBtn"),

    copyPreviewButton:
        document.getElementById("copyPreviewBtn"),

    /* Status */

    offerStatus:
        document.getElementById("offerStatus"),

    workbookStatus:
        document.getElementById("workbookStatus"),

    sheetStatus:
        document.getElementById("sheetStatus"),

    promptBuilderActivity:
        document.getElementById("promptBuilderActivity"),

    workbookDetailsStrip:
        document.getElementById("workbookDetailsStrip"),

    sheetSelectorGroup:
        document.getElementById("sheetSelectorGroup"),

    creditCardMaskStatus:
        document.getElementById("creditCardMaskStatus"),

    ssnMaskStatus:
        document.getElementById("ssnMaskStatus"),

    promptSafetyStatus:
        document.getElementById("promptSafetyStatus"),

    promptSafetyAlert:
        document.getElementById("promptSafetyAlert"),

    generatedPromptStatus:
        document.getElementById("generatedPromptStatus"),

    gptWorkspaceEmpty:
        document.getElementById("gptWorkspaceEmpty"),

    gptIframe:
        document.getElementById("gptIframe"),

    promptReadiness:
        document.getElementById("promptReadiness"),

    temporaryWorkbookAttachment:
        document.getElementById("temporaryWorkbookAttachment"),

    temporaryWorkbookName:
        document.getElementById("temporaryWorkbookName"),

    temporaryWorkbookSheets:
        document.getElementById("temporaryWorkbookSheets"),

    temporaryWorkbookSize:
        document.getElementById("temporaryWorkbookSize"),

    temporaryWorkbookDownload:
        document.getElementById("temporaryWorkbookDownload"),

    maskingSummary:
        document.getElementById("maskingSummary"),

    maskingSummaryCount:
        document.getElementById("maskingSummaryCount"),

    maskingSummaryList:
        document.getElementById("maskingSummaryList"),

    /* Loading */

    loadingOverlay:
        document.getElementById("loadingOverlay"),

    /* Toast */

    liveToast:
        document.getElementById("liveToast")

};

/* ============================================================
   BOOTSTRAP COMPONENTS
============================================================ */

const BootstrapUI = {

    toast: null,

    previewModal: null,

    sheetPreviewModal: null

};

/* ============================================================
   INITIALIZE
============================================================ */

function initializeApplication() {

    initializeBootstrap();

    cacheDefaults();

    registerBaseEvents();

    resetInterface();

    console.log(
        "Workbook Automation Tool Initialized"
    );

}

/* ============================================================
   BOOTSTRAP
============================================================ */

function initializeBootstrap() {

    if (UI.liveToast) {

        BootstrapUI.toast =
            bootstrap.Toast.getOrCreateInstance(
                UI.liveToast
            );

    }

    if (UI.previewModal) {

        BootstrapUI.previewModal =
            bootstrap.Modal.getOrCreateInstance(
                UI.previewModal
            );

    }

}

/* ============================================================
   DEFAULT VALUES
============================================================ */

function cacheDefaults() {

    AppState.workflowStep = 1;

    AppState.loading = false;

    AppState.dirty = false;

}

/* ============================================================
   RESET UI
============================================================ */

function resetInterface() {

    UI.workbookDetailsStrip?.classList.add("d-none");

    UI.sheetSelectorGroup?.classList.add("d-none");

    if (UI.previewSheetButton)
        UI.previewSheetButton.disabled = true;

    updateWorkbookFetchStatus(
        "Waiting",
        "secondary"
    );

    if (UI.workbookName)
        UI.workbookName.textContent = "Not Loaded";

    if (UI.workbookType)
        UI.workbookType.textContent = "-";

    if (UI.workbookSource) {

        UI.workbookSource.textContent = "-";
        UI.workbookSource.className =
            "badge bg-secondary";

    }

    if (UI.sheetPreviewModal) {

        BootstrapUI.sheetPreviewModal =
            bootstrap.Modal.getOrCreateInstance(
                UI.sheetPreviewModal
            );

    }

    if (UI.sheetCount)
        UI.sheetCount.textContent = "0";

    if (UI.aiStatus)
        UI.aiStatus.textContent = "Ready";

    if (UI.piiStatus)
        UI.piiStatus.textContent = "None";

    [
        UI.offer,
        UI.context,
        UI.sheetPrompt,
        UI.finalPrompt
    ].forEach(element => {

        if (!element) return;

        element.value = "";

        element.readOnly = true;

    });

}

/******************************************************************
 * Part 1B
 * Event Registration
 * Workflow
 * Loading
 * Toast
 * Common Helpers
 ******************************************************************/

/* ============================================================
   EVENT REGISTRATION
============================================================ */

function registerBaseEvents() {

    UI.fetchButton?.addEventListener(
        "click",
        fetchWorkbook
    );

    UI.resetWorkspaceButton?.addEventListener(
        "click",
        resetApplication
    );

    UI.previewSheetButton?.addEventListener(
        "click",
        showSheetPreview
    );

    UI.sheetSelector?.addEventListener(
        "change",
        handleSheetSelection
    );

    UI.manualWorkbookInput?.addEventListener(
        "change",
        handleManualWorkbookUpload
    );

    UI.cancelMonitorButton?.addEventListener(
        "click",
        cancelWorkbookMonitoring
    );

    UI.selectAllWorksheetsButton?.addEventListener(
        "click",
        selectAllWorksheets
    );

    UI.clearAllWorksheetsButton?.addEventListener(
        "click",
        clearAllWorksheets
    );

    UI.sheetSelector?.addEventListener(
        "mousedown",
        event => {

            if (
                !UI.sheetSelector.multiple ||
                event.target.tagName !== "OPTION"
            ) {

                return;

            }

            event.preventDefault();
            event.target.selected =
                !event.target.selected;
            handleSheetSelection();

        }
    );
}

/* ============================================================
   WORKFLOW
============================================================ */

function setWorkflowStep(step, pending = false) {

    AppState.workflowStep = step;

    const labels = {
        1: pending ? "Downloading Workbook" : "⏳ Waiting for Workbook",
        2: "Loading Workbook",
        3: "Loading Workbook",
        4: "Building Prompt",
        5: "✅ Prompt Ready",
        6: "✅ Prompt Ready"
    };

    updatePromptBuilderActivity(
        labels[step] || "⏳ Waiting for Workbook",
        pending ? "warning" :
            step >= 5 ? "success" : "secondary",
        pending
    );

}

function updatePromptBuilderActivity(
    message,
    color = "secondary",
    pending = false
) {

    if (!UI.promptBuilderActivity) return;

    UI.promptBuilderActivity.className =
        `prompt-builder-activity text-${color} align-self-center`;

    UI.promptBuilderActivity.replaceChildren();

    if (pending) {

        const spinner = document.createElement("i");
        spinner.className =
            "bi bi-arrow-repeat status-spinner me-1";
        spinner.setAttribute("aria-hidden", "true");
        UI.promptBuilderActivity.appendChild(spinner);

    }

    UI.promptBuilderActivity.appendChild(
        document.createTextNode(message)
    );

}

/* ============================================================
   LOADING
============================================================ */

function showLoading(showOverlay = true) {

    AppState.loading = true;

    if (showOverlay) {

        UI.loadingOverlay?.classList.remove("d-none");

    }

}

function hideLoading() {

    AppState.loading = false;

    UI.loadingOverlay?.classList.add("d-none");

}

/* ============================================================
   BUTTON STATE
============================================================ */

function disableFetchButton() {

    if (UI.fetchButton) {

        UI.fetchButton.disabled = true;

        UI.fetchButton.innerHTML = `
            <span class="spinner-border spinner-border-sm me-2"></span>
            Fetching...
        `;

    }

    updateWorkbookFetchStatus(
        "Loading",
        "warning",
        true
    );

}

function enableFetchButton() {

    if (UI.fetchButton) {

        UI.fetchButton.disabled = false;

        UI.fetchButton.innerHTML = `
            <i class="bi bi-cloud-download me-2"></i>
            Fetch Workbook
        `;

    }

}

function updateWorkbookFetchStatus(
    message,
    color,
    pending = false
) {

    if (!UI.workbookFetchStatus) return;

    UI.workbookFetchStatus.className =
        `workbook-fetch-status text-${color}`;

    UI.workbookFetchStatus.replaceChildren();

    if (pending) {

        const spinner = document.createElement("i");
        spinner.className =
            "bi bi-arrow-repeat status-spinner me-1";
        UI.workbookFetchStatus.appendChild(spinner);

    }

    else {

        const icon =
            color === "success" ? "🟢 " :
                color === "danger" ? "🔴 " : "⚪ ";

        UI.workbookFetchStatus.appendChild(
            document.createTextNode(icon)
        );

    }

    UI.workbookFetchStatus.appendChild(
        document.createTextNode(message)
    );

}

/* ============================================================
   TOAST
============================================================ */

function showToast(message) {

    if (!BootstrapUI.toast) {

        alert(message);

        return;

    }

    const body =
        UI.liveToast.querySelector(".toast-body");

    if (body) {

        body.innerHTML = `
            <i class="bi bi-check-circle-fill me-2"></i>
            ${message}
        `;

    }

    BootstrapUI.toast.show();

}

/* ============================================================
   STATUS BADGES
============================================================ */

function setBadge(element, text, color, pending = false) {

    if (!element) return;

    element.className = `badge bg-${color}`;
    element.replaceChildren();

    if (pending) {

        const spinner = document.createElement("i");
        spinner.className =
            "bi bi-arrow-repeat status-spinner me-1";
        spinner.setAttribute("aria-hidden", "true");
        element.appendChild(spinner);

    }

    element.appendChild(
        document.createTextNode(text)
    );

}

function updatePromptStatus(section, state) {

    const map = {

        offer: UI.offerStatus,

        workbook: UI.workbookStatus,

        sheet: UI.sheetStatus

    };

    const badge = map[section];

    if (!badge) return;

    const iconStates = {
        ready: ["success", "Ready"],
        editing: ["warning", "Editing"],
        waiting: ["secondary", "Waiting"],
        pending: ["warning", "Loading"],
        error: ["danger", "Error"]
    };

    const [iconColor, iconLabel] =
        iconStates[state] || iconStates.waiting;

    badge.className =
        `section-status text-${iconColor}`;
    badge.textContent = "●";
    badge.title = iconLabel;
    badge.setAttribute("aria-label", iconLabel);

    return;

    switch (state) {

        case "ready":

            setBadge(
                badge,
                "🟢 Ready",
                "success"
            );

            break;

        case "editing":

            setBadge(
                badge,
                "Editing",
                "warning"
            );

            break;

        case "waiting":

            setBadge(
                badge,
                "🟡 Waiting",
                "warning"
            );

            break;

        case "pending":

            setBadge(
                badge,
                "Pending",
                "primary",
                true
            );

            break;

        default:

            setBadge(
                badge,
                state,
                "primary"
            );

    }

}

/* ============================================================
   AI STATUS
============================================================ */

function updateAIStatus(message) {

    if (!UI.aiStatus) return;

    const normalized =
        String(message).toLowerCase();

    const pending =
        normalized.endsWith("...") ||
        /(searching|scanning|loading|processing|downloading)/.test(normalized);

    const complete =
        /(ready|downloaded|loaded|copied|complete|safe)/.test(normalized);

    const failed =
        /(error|failed|unsafe)/.test(normalized);

    setBadge(
        UI.aiStatus,
        message,
        failed ? "danger" :
            complete ? "success" :
                pending ? "warning" : "secondary",
        pending
    );

}

/* ============================================================
   PII STATUS
============================================================ */

function updatePIIBadge(hasPII, count = 0) {

    if (!UI.piiStatus) return;

    if (hasPII) {

        UI.piiStatus.className =
            "badge bg-danger";

        UI.piiStatus.textContent =
            `${count} Found`;

    }

    else {

        UI.piiStatus.className =
            "badge bg-success";

        UI.piiStatus.textContent =
            "None";

    }

}

/* ============================================================
   SAFE VALUE
============================================================ */

function safeText(value, fallback = "-") {

    if (
        value === undefined ||
        value === null ||
        value === ""
    ) {

        return fallback;

    }

    return value;

}

/******************************************************************
 * Part 1C
 * Clipboard
 * Prompt Preview
 * Local Storage
 * Utility Functions
 * Reset Helpers
 ******************************************************************/

/* ============================================================
   CLIPBOARD
============================================================ */

async function copyToClipboard(text, successMessage = "Copied successfully.") {

    try {

        if (!text) {

            showToast("Nothing to copy.");

            return false;

        }

        await navigator.clipboard.writeText(text);

        showToast(successMessage);

        return true;

    }

    catch (error) {

        console.error(error);

        alert("Unable to copy.");

        return false;

    }

}

/* ============================================================
   PREVIEW MODAL
============================================================ */

function updatePreview() {

    if (!UI.previewArea || !UI.finalPrompt)
        return;

    UI.previewArea.value =
        UI.finalPrompt.value;

}

function openPreview() {

    updatePreview();

    BootstrapUI.previewModal?.show();

}

function closePreview() {

    BootstrapUI.previewModal?.hide();

}

/* ============================================================
   CHARACTER COUNT
============================================================ */

function characterCount(text) {

    return text
        ? text.length
        : 0;

}

function wordCount(text) {

    if (!text) return 0;

    return text
        .trim()
        .split(/\s+/)
        .filter(Boolean)
        .length;

}

/* ============================================================
   LOCAL STORAGE
============================================================ */

const STORAGE_KEY =
    "WorkbookAutomationDraft";

function saveDraft() {

    try {

        const draft = {

            offer:
                UI.offer?.value || "",

            context:
                UI.context?.value || "",

            sheet:
                UI.sheetPrompt?.value || "",

            prompt:
                UI.finalPrompt?.value || "",

            selectedSheet:
                AppState.selectedSheet,

            selectedSheets:
                AppState.selectedSheets,

            savedOn:
                new Date().toISOString()

        };

        localStorage.setItem(
            STORAGE_KEY,
            JSON.stringify(draft)
        );

    }

    catch (error) {

        console.warn(error);

    }

}

function loadDraft() {

    try {

        const draft =
            JSON.parse(
                localStorage.getItem(STORAGE_KEY)
            );

        if (!draft) return;

        if (UI.offer)
            UI.offer.value =
                draft.offer || "";

        if (UI.context)
            UI.context.value =
                draft.context || "";

        if (UI.sheetPrompt)
            UI.sheetPrompt.value =
                draft.sheet || "";

        if (UI.finalPrompt)
            UI.finalPrompt.value =
                draft.prompt || "";

    }

    catch (error) {

        console.warn(error);

    }

}

function clearDraft() {

    localStorage.removeItem(
        STORAGE_KEY
    );

}

/* ============================================================
   RESET PROMPT BUILDER
============================================================ */

function clearPromptBuilder() {

    [

        UI.offer,

        UI.context,

        UI.sheetPrompt,

        UI.finalPrompt

    ].forEach(element => {

        if (element)
            element.value = "";

    });

    updatePromptStatus(
        "offer",
        "waiting"
    );

    updatePromptStatus(
        "workbook",
        "waiting"
    );

    updatePromptStatus(
        "sheet",
        "waiting"
    );

}

/* ============================================================
   RESET SHEET LIST
============================================================ */

function resetSheetSelector() {

    if (!UI.sheetSelector)
        return;

    UI.sheetSelector.replaceChildren();
    UI.worksheetCheckboxList?.replaceChildren();
    if (UI.worksheetDropdownButton)
        UI.worksheetDropdownButton.textContent =
            "Select worksheets";

}

/* ============================================================
   DATE
============================================================ */

function today() {

    return new Date()
        .toLocaleDateString();

}

/* ============================================================
   DELAY
============================================================ */

function delay(ms) {

    return new Promise(resolve => {

        setTimeout(resolve, ms);

    });

}

/* ============================================================
   DEBUG LOGGER
============================================================ */

function log(...message) {

    console.log(
        "[Workbook Tool]",
        ...message
    );

}

/* ============================================================
   INITIAL STATE
============================================================ */

setWorkflowStep(1);

updatePromptStatus(
    "offer",
    "waiting"
);

updatePromptStatus(
    "workbook",
    "waiting"
);

updatePromptStatus(
    "sheet",
    "waiting"
);

loadDraft();

log("Globals Loaded");

/******************************************************************
 * END OF 01_globals.js
 ******************************************************************/

/******************************************************************
 * 02_fetchWorkbook.js
 * Part 2A
 * Workbook Fetch API
 ******************************************************************/

/* ============================================================
   FETCH WORKBOOK
============================================================ */

async function fetchWorkbook() {

    const projectReference =
        UI.projectRef?.value.trim();

    const taskNumber =
        UI.taskNumber?.value.trim();

    if (!projectReference) {

        showToast("Please enter Project Reference.");

        UI.projectRef?.focus();

        return;

    }

    if (!taskNumber) {

        showToast("Please enter Task Number.");

        UI.taskNumber?.focus();

        return;

    }

    try {

        endWorkbookSession();
        disableFetchButton();

        showLoading();

        clearGeneratedPrompt();

        resetPromptSafety();

        UI.workbookDetailsStrip?.classList.add("d-none");

        UI.sheetSelectorGroup?.classList.add("d-none");

        setWorkflowStep(1, true);

        updatePromptStatus("offer", "pending");

        updatePromptStatus("workbook", "pending");

        updatePromptStatus("sheet", "pending");

        updateAIStatus("Searching Project...");

        updatePromptBuilderActivity(
            "Searching for workbook...",
            "warning",
            true
        );

        updateWorkbookFetchStatus(
            "Searching",
            "warning",
            true
        );

        UI.manualDownloadPanel?.classList.add(
            "d-none"
        );

        const lookupResponse = await fetch(
            "/lookup_workbook",
            {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    project_reference: projectReference,
                    task_number: taskNumber
                })
            }
        );

        const lookup =
            await lookupResponse.json()
                .catch(() => ({}));

        if (!lookupResponse.ok) {

            throw new Error(
                lookup.error ||
                "Unable to search for the workbook."
            );

        }

        if (lookup.status === "manual_required") {

            showManualDownload(
                lookup.message,
                lookup.task_url,
                lookup.filename
            );

            hideLoading();
            enableFetchButton();
            return;

        }

        updatePromptBuilderActivity(
            "Downloading workbook...",
            "warning",
            true
        );

        updateWorkbookFetchStatus(
            "Downloading",
            "warning",
            true
        );

        const response = await fetch(
            "/fetch_workbook",
            {

                method: "POST",

                headers: {

                    "Content-Type":
                        "application/json"

                },

                body: JSON.stringify({

                    project_reference:
                        projectReference,

                    task_number:
                        taskNumber

                })

            }
        );

        if (!response.ok) {

            const error =
                await response.json()
                    .catch(() => ({}));

            throw new Error(

                error.error ||

                "Unable to fetch workbook."

            );

        }

        const data =
            await response.json();

        if (data.status === "manual_required") {

            showManualDownload(
                data.message,
                data.task_url,
                data.expected_filename
            );

            hideLoading();
            enableFetchButton();
            return;

        }

        updatePromptBuilderActivity(
            "Loading workbook...",
            "warning",
            true
        );

        updateWorkbookFetchStatus(
            "Loading",
            "warning",
            true
        );

        AppState.workbook =
            data;

        setWorkflowStep(2, true);

        updateAIStatus(
            "Workbook Downloaded"
        );

        loadWorkbook(data);

    }

    catch (error) {

        console.error(error);

        setWorkflowStep(1);

        updatePromptBuilderActivity(
            "Error Loading Workbook",
            "danger"
        );

        updateWorkbookFetchStatus(
            "Error",
            "danger"
        );

        updatePromptStatus("offer", "waiting");

        updatePromptStatus("workbook", "waiting");

        updatePromptStatus("sheet", "waiting");

        updateAIStatus("Error");

        hideLoading();

        enableFetchButton();

        showToast(error.message);

    }

}

function showManualDownload(
    message,
    taskUrl,
    expectedFilename
) {

    if (UI.manualDownloadMessage) {

        UI.manualDownloadMessage.textContent =
            message ||
            "Workbook download link not found. Please download the workbook manually using the task link below.";

    }

    if (UI.expectedWorkbookName) {

        UI.expectedWorkbookName.textContent =
            expectedFilename
                ? `Expected workbook: ${expectedFilename}`
                : "";

        UI.expectedWorkbookName.classList.toggle(
            "d-none",
            !expectedFilename
        );

    }

    if (UI.openTaskButton) {

        UI.openTaskButton.href =
            taskUrl || "#";
        UI.openTaskButton.classList.toggle(
            "d-none",
            !taskUrl
        );

    }

    UI.manualDownloadPanel?.classList.remove(
        "d-none"
    );

    updatePromptBuilderActivity(
        "Manual download required",
        "danger"
    );

    updateWorkbookFetchStatus(
        "Manual",
        "danger"
    );

    updateAIStatus("Manual Download Required");

    UI.manualUploadLabel?.classList.add("d-none");
    startWorkbookMonitoring(expectedFilename);

}

async function startWorkbookMonitoring(expectedFilename) {

    await cancelWorkbookMonitoring(false);

    updateMonitorStatus("Monitoring Downloads folder...", "warning", true);

    try {

        const response = await fetch(
            "/monitor_workbook/start",
            {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    project_reference:
                        UI.projectRef?.value.trim() || "",
                    task_number:
                        UI.taskNumber?.value.trim() || "",
                    expected_filename:
                        expectedFilename || ""
                })
            }
        );
        const result = await response.json();
        if (!response.ok)
            throw new Error(result.error || "Unable to start folder monitoring.");

        AppState.monitorId = result.monitor_id;
        updateMonitorStatus(
            `Monitoring Downloads folder: ${result.directory}`,
            "warning",
            true
        );
        AppState.monitorTimer = window.setInterval(
            pollWorkbookMonitoring,
            1500
        );

    }
    catch (error) {

        updateMonitorStatus(error.message, "danger");
        UI.manualUploadLabel?.classList.remove("d-none");

    }

}

async function pollWorkbookMonitoring() {

    if (!AppState.monitorId) return;

    try {

        const response = await fetch(
            "/monitor_workbook/poll",
            {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    monitor_id: AppState.monitorId
                })
            }
        );
        const result = await response.json();
        if (!response.ok)
            throw new Error(result.error || "Folder monitoring failed.");

        updateMonitorStatus(
            result.message || "Waiting for workbook download...",
            result.status === "ready" ? "success" : "warning",
            result.status === "monitoring"
        );

        if (result.status === "ready") {

            stopMonitorTimer();
            AppState.workbook = result;
            updatePromptBuilderActivity(
                "Workbook uploaded successfully.",
                "success"
            );
            updateAIStatus("Workbook Uploaded");
            loadWorkbook(result);

        }
        else if (
            result.status === "timeout" ||
            result.status === "cancelled"
        ) {

            stopMonitorTimer();
            UI.manualUploadLabel?.classList.remove("d-none");

        }

    }
    catch (error) {

        stopMonitorTimer();
        updateMonitorStatus(error.message, "danger");
        UI.manualUploadLabel?.classList.remove("d-none");

    }

}

async function cancelWorkbookMonitoring(notifyServer = true) {

    const monitorId = AppState.monitorId;
    stopMonitorTimer();

    if (notifyServer && monitorId) {

        await fetch(
            "/monitor_workbook/cancel",
            {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    monitor_id: monitorId
                })
            }
        ).catch(() => {});
        updateMonitorStatus(
            "Workbook monitoring cancelled.",
            "secondary"
        );
        UI.manualUploadLabel?.classList.remove("d-none");

    }

}

function stopMonitorTimer() {

    if (AppState.monitorTimer)
        window.clearInterval(AppState.monitorTimer);
    AppState.monitorTimer = null;
    AppState.monitorId = "";

}

function updateMonitorStatus(message, color, pending = false) {

    if (!UI.downloadMonitorStatus) return;
    UI.downloadMonitorStatus.className =
        `download-monitor-status mt-2 text-${color}`;
    UI.downloadMonitorStatus.replaceChildren();
    if (pending) {

        const spinner = document.createElement("span");
        spinner.className =
            "spinner-border spinner-border-sm me-2";
        UI.downloadMonitorStatus.appendChild(spinner);

    }
    UI.downloadMonitorStatus.appendChild(
        document.createTextNode(message)
    );

}

async function handleManualWorkbookUpload(event) {

    const file = event.target.files?.[0];
    if (!file) return;

    await cancelWorkbookMonitoring();

    const form = new FormData();
    form.append("workbook", file);
    form.append(
        "project_reference",
        UI.projectRef?.value.trim() || ""
    );
    form.append(
        "task_number",
        UI.taskNumber?.value.trim() || ""
    );

    try {

        updatePromptBuilderActivity(
            "Workbook detected and uploading...",
            "warning",
            true
        );

        updateWorkbookFetchStatus(
            "Uploading",
            "warning",
            true
        );

        const response = await fetch(
            "/upload_workbook",
            {
                method: "POST",
                body: form
            }
        );

        const data =
            await response.json()
                .catch(() => ({}));

        if (!response.ok) {

            throw new Error(
                data.error ||
                "Unable to upload the workbook."
            );

        }

        updatePromptBuilderActivity(
            "Loading workbook...",
            "warning",
            true
        );

        updateWorkbookFetchStatus(
            "Loading",
            "warning",
            true
        );

        AppState.workbook = data;
        UI.manualDownloadPanel?.classList.add(
            "d-none"
        );
        event.target.value = "";
        loadWorkbook(data);
        showToast(
            "Workbook detected and uploaded."
        );

    }

    catch (error) {

        console.error(error);
        updatePromptBuilderActivity(
            "Workbook upload failed",
            "danger"
        );
        updateWorkbookFetchStatus(
            "Error",
            "danger"
        );
        showToast(error.message);

    }

}

/* ============================================================
   LOAD WORKBOOK
============================================================ */

function loadWorkbook(data) {

    AppState.workbookName =
        data.workbook_name || "";

    AppState.workbookType =
        data.workbook_type || "";

    UI.workbookName.textContent =
        safeText(data.workbook_name);

    UI.workbookType.textContent =
        safeText(data.workbook_type);

    if (UI.workbookSource) {

        const manual =
            data.upload_source === "manual";

        UI.workbookSource.className =
            `badge bg-${manual ? "primary" : "success"}`;

        UI.workbookSource.textContent =
            manual
                ? "Manual Upload 📤"
                : "Auto Downloaded ✅";

        UI.workbookSource.title =
            manual
                ? "Manual Upload"
                : "Auto Downloaded";

    }

    UI.sheetCount.textContent =
        data.sheet_count || 0;

    updatePIIBadge(

        data.has_pii || false,

        data.pii_count || 0

    );

    populateSheets(

        data.sheets || []

    );

    UI.workbookDetailsStrip?.classList.remove("d-none");

    UI.sheetSelectorGroup?.classList.remove("d-none");

    if (UI.previewSheetButton)
        UI.previewSheetButton.disabled = false;

    enablePromptBuilder();

    setWorkflowStep(3, true);

    loadDefaultSheet();

    updateAIStatus(

        "Workbook Loaded"

    );

    hideLoading();

    enableFetchButton();

    updateWorkbookFetchStatus(
        "Ready",
        "success"
    );

    showToast(

        "Workbook loaded successfully."

    );

}

/******************************************************************
 * Part 2B
 * Sheet Loader
 * Prompt Builder Population
 ******************************************************************/

/* ============================================================
   POPULATE SHEET DROPDOWN
============================================================ */

function populateSheets(sheetList) {

    resetSheetSelector();
    AppState.worksheetSelectionOrder = [];
    UI.worksheetCheckboxList?.replaceChildren();

    if (!Array.isArray(sheetList))
        return;

    sheetList.forEach(sheet => {

        const option =
            document.createElement("option");

        option.value = sheet;

        option.textContent = sheet;

        UI.sheetSelector.appendChild(option);

        const item = document.createElement("label");
        item.className = "worksheet-checkbox-item";

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.className = "form-check-input";
        checkbox.value = sheet;
        checkbox.addEventListener("change", () => {

            option.selected = checkbox.checked;
            handleSheetSelection();

        });

        const name = document.createElement("span");
        name.textContent = sheet;
        item.append(checkbox, name);
        UI.worksheetCheckboxList?.appendChild(item);

    });

}

function syncWorksheetDropdown() {

    const selected = AppState.selectedSheets;

    if (UI.worksheetDropdownButton) {

        UI.worksheetDropdownButton.textContent =
            selected.length === 0
                ? "Select worksheets"
                : selected.length <= 2
                    ? selected.join(", ")
                    : `${selected.length} Sheets Selected`;

    }

    UI.worksheetCheckboxList
        ?.querySelectorAll('input[type="checkbox"]')
        .forEach(checkbox => {

            checkbox.checked =
                selected.includes(checkbox.value);

        });

}

function selectAllWorksheets() {

    Array.from(UI.sheetSelector.options)
        .forEach(option => {

            option.selected = true;

        });
    AppState.worksheetSelectionOrder =
        Array.from(UI.sheetSelector.options)
            .map(option => option.value);
    handleSheetSelection();

}

function clearAllWorksheets() {

    Array.from(UI.sheetSelector.options)
        .forEach(option => {

            option.selected = false;

        });
    AppState.worksheetSelectionOrder = [];
    handleSheetSelection();

}

/* ============================================================
   SHEET CHANGED
============================================================ */

function handleSheetSelection() {

    const currentlySelected =
        Array.from(
            UI.sheetSelector.selectedOptions
        ).map(option => option.value);

    const selectedSet =
        new Set(currentlySelected);

    const retainedSelectionOrder =
        AppState.worksheetSelectionOrder.filter(
            name => selectedSet.has(name)
        );

    currentlySelected.forEach(name => {

        if (!retainedSelectionOrder.includes(name))
            retainedSelectionOrder.push(name);

    });

    AppState.worksheetSelectionOrder =
        retainedSelectionOrder;

    const sheetNames =
        [...AppState.worksheetSelectionOrder];

    if (!sheetNames.length) {

        AppState.selectedSheet = "";
        AppState.selectedSheets = [];
        AppState.worksheetSelectionOrder = [];
        UI.worksheetSelectionMessage?.classList.remove(
            "d-none"
        );
        if (UI.previewSheetButton)
            UI.previewSheetButton.disabled = true;
        syncWorksheetDropdown();
        renderWorkbookContent([]);
        clearTemporaryWorkbookAttachment();
        clearPromptBuilder();
        clearGeneratedPrompt();
        updatePromptBuilderActivity(
            "Select at least one worksheet",
            "secondary"
        );
        return;

    }

    UI.worksheetSelectionMessage?.classList.add("d-none");
    if (UI.previewSheetButton)
        UI.previewSheetButton.disabled = false;

    AppState.selectedSheet =
        sheetNames[0];

    AppState.selectedSheets =
        sheetNames;
    syncWorksheetDropdown();
    renderWorkbookContent(sheetNames);

    setWorkflowStep(4, true);

    updateAIStatus(
        "Loading Sheet..."
    );

    loadSelectedSheets(sheetNames);

}

function renderWorkbookContent(sheetNames) {

    if (!UI.workbookContentList) return;
    UI.workbookContentList.replaceChildren();

    sheetNames.forEach(name => {

        const row = document.createElement("div");
        row.className = "workbook-content-item";

        const label = document.createElement("span");
        label.innerHTML =
            '<i class="bi bi-file-earmark-spreadsheet me-2"></i>';
        label.appendChild(document.createTextNode(name));

        const preview = document.createElement("button");
        preview.type = "button";
        preview.className = "btn btn-outline-primary btn-sm";
        preview.innerHTML =
            '<i class="bi bi-eye me-1"></i>Preview';
        preview.addEventListener(
            "click",
            () => {

                renderSheetPreview(name);
                BootstrapUI.sheetPreviewModal?.show();

            }
        );

        row.append(label, preview);
        UI.workbookContentList.appendChild(row);

    });

}

function showSheetPreview() {

    if (!AppState.workbook || !AppState.selectedSheet) {

        showToast("Select a worksheet to preview.");
        return;

    }

    renderSheetPreview(AppState.selectedSheet);
    BootstrapUI.sheetPreviewModal?.show();

}

function renderSheetPreview(sheetName) {

    const sheet =
        AppState.workbook?.sheet_data?.[sheetName];

    if (!sheet || !UI.sheetPreviewGrid) return;

    UI.sheetPreviewWorkbookName.textContent =
        AppState.workbookName || "";
    UI.sheetPreviewGrid.replaceChildren();

    const rows =
        Array.isArray(sheet.preview_rows)
            ? sheet.preview_rows
            : [];
    const columnCount = rows.reduce(
        (maximum, row) =>
            Math.max(maximum, Array.isArray(row) ? row.length : 0),
        0
    );

    const table = document.createElement("table");
    table.className = "sheet-preview-table";

    const head = document.createElement("thead");
    const headingRow = document.createElement("tr");
    const corner = document.createElement("th");
    corner.className = "sheet-row-heading sheet-corner";
    headingRow.appendChild(corner);

    for (let column = 0; column < columnCount; column += 1) {

        const heading = document.createElement("th");
        heading.textContent = excelColumnName(column);
        headingRow.appendChild(heading);

    }

    head.appendChild(headingRow);
    table.appendChild(head);

    const body = document.createElement("tbody");

    rows.forEach((row, rowIndex) => {

        const tableRow = document.createElement("tr");
        const rowHeading = document.createElement("th");
        rowHeading.className = "sheet-row-heading";
        rowHeading.textContent = String(rowIndex + 1);
        tableRow.appendChild(rowHeading);

        for (let column = 0; column < columnCount; column += 1) {

            const cell = document.createElement("td");
            cell.textContent = row[column] ?? "";
            cell.title = row[column] ?? "";
            tableRow.appendChild(cell);

        }

        body.appendChild(tableRow);

    });

    table.appendChild(body);
    UI.sheetPreviewGrid.appendChild(table);

    UI.sheetPreviewNotice?.classList.toggle(
        "d-none",
        !sheet.preview_truncated
    );
    if (sheet.preview_truncated && UI.sheetPreviewNotice) {

        UI.sheetPreviewNotice.textContent =
            "Preview is limited to the first 5,000 rows and 200 columns.";

    }

    renderSheetPreviewTabs(sheetName);

}

function renderSheetPreviewTabs(activeSheet) {

    if (!UI.sheetPreviewTabs) return;
    UI.sheetPreviewTabs.replaceChildren();

    (AppState.workbook?.sheets || []).forEach(sheetName => {

        const tab = document.createElement("button");
        tab.type = "button";
        tab.className =
            `sheet-preview-tab${sheetName === activeSheet ? " active" : ""}`;
        tab.textContent = sheetName;
        tab.setAttribute(
            "aria-selected",
            String(sheetName === activeSheet)
        );
        tab.addEventListener("click", () => {

            renderSheetPreview(sheetName);

        });
        UI.sheetPreviewTabs.appendChild(tab);

    });

}

function excelColumnName(index) {

    let value = index + 1;
    let name = "";

    while (value > 0) {

        value -= 1;
        name =
            String.fromCharCode(65 + (value % 26)) + name;
        value = Math.floor(value / 26);

    }

    return name;

}

async function loadSelectedSheets(sheetNames) {

    const selected = sheetNames
        .map(name => ({
            name,
            data: AppState.workbook
                ?.sheet_data?.[name]
        }))
        .filter(item => item.data);

    if (!selected.length) {

        showToast("Unable to load the selected worksheets.");
        return;

    }

    const offer =
        selected.find(item => item.data.offer_description)
            ?.data.offer_description || "";

    const context =
        `Selected worksheets: ${sheetNames.join(", ")}`;

    const instructions = selected
        .map(item =>
            `### ${item.name}\n${item.data.sheet_prompt || "No configured rules."}`
        )
        .join("\n\n");

    AppState.prompt.offer = offer;
    AppState.prompt.context = context;
    AppState.prompt.sheet = instructions;

    UI.offer.value = offer;
    UI.context.value = context;
    UI.sheetPrompt.value = instructions;

    updatePromptStatus("offer", offer ? "ready" : "error");
    updatePromptStatus("workbook", context ? "ready" : "error");
    updatePromptStatus("sheet", instructions ? "ready" : "error");

    const attachmentReady =
        await createTemporaryWorkbook(sheetNames);

    if (!attachmentReady) {

        clearGeneratedPrompt();
        updatePromptBuilderActivity(
            "Unable to prepare temporary workbook",
            "danger"
        );
        return;

    }

    regeneratePrompt();
    saveDraft();

    if (hasCompletePromptData()) {

        setWorkflowStep(5);

    }

    updateAIStatus("Prompt Builder Ready");
    showToast(`${selected.length} worksheet${selected.length === 1 ? "" : "s"} loaded`);

}

async function createTemporaryWorkbook(sheetNames) {

    const sessionId =
        AppState.workbook?.workbook_session_id;

    if (!sessionId || !sheetNames.length)
        return false;

    updatePromptBuilderActivity(
        "Building temporary workbook...",
        "warning",
        true
    );

    try {

        const response = await fetch(
            "/temporary_workbook",
            {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    workbook_session_id: sessionId,
                    selected_sheets: sheetNames
                })
            }
        );
        const attachment =
            await response.json();

        if (!response.ok)
            throw new Error(
                attachment.error ||
                "Unable to create temporary workbook."
            );

        AppState.temporaryWorkbook = attachment;
        renderTemporaryWorkbookAttachment();
        return true;

    }
    catch (error) {

        console.error(error);
        showToast(error.message);
        clearTemporaryWorkbookAttachment();
        return false;

    }

}

function renderTemporaryWorkbookAttachment() {

    const attachment =
        AppState.temporaryWorkbook;

    if (!attachment) {

        clearTemporaryWorkbookAttachment();
        return;

    }

    UI.temporaryWorkbookName.textContent =
        attachment.filename;
    UI.temporaryWorkbookSheets.textContent =
        attachment.sheets.join(", ");
    UI.temporaryWorkbookSize.textContent =
        formatFileSize(attachment.size);
    UI.temporaryWorkbookDownload.href =
        attachment.download_url;
    UI.temporaryWorkbookAttachment?.classList.remove(
        "d-none"
    );

}

function clearTemporaryWorkbookAttachment() {

    AppState.temporaryWorkbook = null;
    UI.temporaryWorkbookAttachment?.classList.add(
        "d-none"
    );
    if (UI.temporaryWorkbookDownload)
        UI.temporaryWorkbookDownload.href = "#";

}

function formatFileSize(bytes) {

    if (!Number.isFinite(Number(bytes)))
        return "";
    if (bytes < 1024)
        return `${bytes} B`;
    if (bytes < 1024 * 1024)
        return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;

}

/* ============================================================
   LOAD SHEET DATA
============================================================ */

function loadSelectedSheet(sheetName) {

    return loadSelectedSheets([sheetName]);

    const sheet =
        AppState.workbook
            ?.sheet_data?.[sheetName];

    if (!sheet) {

        showToast(
            "Unable to load worksheet."
        );

        return;

    }

    /* Store */

    AppState.prompt.offer =
        sheet.offer_description || "";

    AppState.prompt.context =
        sheet.workbook_context || "";

    AppState.prompt.sheet =
        sheet.sheet_prompt || "";

    AppState.prompt.final =
        sheet.generated_prompt || "";

    /* Populate UI */

    UI.offer.value =
        AppState.prompt.offer;

    UI.context.value =
        AppState.prompt.context;

    UI.sheetPrompt.value =
        AppState.prompt.sheet;

    UI.finalPrompt.value =
        AppState.prompt.final;

    /* Update Status */

    updatePromptStatus(
        "offer",
        "ready"
    );

    updatePromptStatus(
        "workbook",
        "ready"
    );

    updatePromptStatus(
        "sheet",
        "ready"
    );

    updatePreview();

    saveDraft();

    if (hasCompletePromptData()) {

        setWorkflowStep(5);

    }

    else {

        updatePromptBuilderActivity(
            "Prompt Data Incomplete",
            "danger"
        );

    }

    updateAIStatus(
        "Prompt Builder Ready"
    );

    showToast(
        `${sheetName} loaded`
    );

}

/* ============================================================
   LOAD FIRST SHEET
============================================================ */

function loadDefaultSheet() {

    if (!UI.sheetSelector)
        return;

    if (
        UI.sheetSelector.options.length === 0
    )
        return;

    UI.sheetSelector.options[0].selected = true;

    handleSheetSelection();

}

/* ============================================================
   REFRESH CURRENT SHEET
============================================================ */

function refreshCurrentSheet() {

    if (!AppState.selectedSheets.length)
        return;

    loadSelectedSheets(
        AppState.selectedSheets
    );

}

/* ============================================================
   CLEAR WORKBOOK
============================================================ */

function clearWorkbook() {

    AppState.workbook = null;

    AppState.selectedSheet = "";

    AppState.selectedSheets = [];
    AppState.worksheetSelectionOrder = [];

    AppState.workbookName = "";

    AppState.workbookType = "";

    UI.workbookDetailsStrip?.classList.add("d-none");

    UI.sheetSelectorGroup?.classList.add("d-none");

    if (UI.workbookName)
        UI.workbookName.textContent =
            "Not Loaded";

    if (UI.workbookSource) {

        UI.workbookSource.textContent = "-";
        UI.workbookSource.className =
            "badge bg-secondary";

    }

    if (UI.workbookType)
        UI.workbookType.textContent =
            "-";

    if (UI.sheetCount)
        UI.sheetCount.textContent =
            "0";

    resetSheetSelector();

    clearPromptBuilder();
    renderWorkbookContent([]);
    clearTemporaryWorkbookAttachment();

    updateAIStatus(
        "Ready"
    );

    updatePIIBadge(false);

    setWorkflowStep(1);

}

/******************************************************************
 * 03_promptBuilder.js
 * Part 3A
 * Prompt Builder
 * Edit Mode
 ******************************************************************/

"use strict";

/* ============================================================
   EDIT STATE
============================================================ */

const EditState = {

    offer: false,

    workbook: false,

    sheet: false

};

/* ============================================================
   REGISTER EDIT EVENTS
============================================================ */

function registerPromptBuilderEvents() {

    document
        .getElementById("editOfferBtn")
        ?.addEventListener(
            "click",
            () => toggleEditor("offer")
        );

    document
        .getElementById("editWorkbookBtn")
        ?.addEventListener(
            "click",
            () => toggleEditor("workbook")
        );

    document
        .getElementById("editSheetBtn")
        ?.addEventListener(
            "click",
            () => toggleEditor("sheet")
        );

}

/* ============================================================
   TOGGLE EDITOR
============================================================ */

function toggleEditor(section) {

    switch (section) {

        case "offer":

            toggleTextarea(

                UI.offer,

                "offer",

                "editOfferBtn"

            );

            break;

        case "workbook":

            toggleTextarea(

                UI.context,

                "workbook",

                "editWorkbookBtn"

            );

            break;

        case "sheet":

            toggleTextarea(

                UI.sheetPrompt,

                "sheet",

                "editSheetBtn"

            );

            break;

    }

}

/* ============================================================
   GENERIC TOGGLE
============================================================ */

function toggleTextarea(

    textarea,

    key,

    buttonId

) {

    if (!textarea) return;

    const button =

        document.getElementById(buttonId);

    EditState[key] =

        !EditState[key];

    textarea.readOnly =

        !EditState[key];

    textarea.classList.toggle(

        "border-primary",

        EditState[key]

    );

    if (EditState[key]) {

        button.innerHTML = `

            <i class="bi bi-check-lg"></i>

            Save

        `;

        textarea.focus();

        updatePromptStatus(

            key,

            "editing"

        );

        expandSection(textarea);

    }

    else {

        button.innerHTML = `

            <i class="bi bi-pencil"></i>

            Edit

        `;

        updatePromptStatus(

            key,

            "ready"

        );

        regeneratePrompt();

        saveDraft();

        showToast(

            "Changes saved."

        );

    }

}

/******************************************************************
 * Part 3B
 * Auto Save
 * Dirty Tracking
 * Accordion
 * Live Prompt Refresh
 ******************************************************************/

/* ============================================================
   AUTO SAVE TIMERS
============================================================ */

const AutoSave = {

    timers: {}

};

/* ============================================================
   REGISTER INPUT EVENTS
============================================================ */

function registerPromptInputEvents() {

    [
        UI.offer,
        UI.context,
        UI.sheetPrompt
    ].forEach(element => {

        if (!element) return;

        element.addEventListener(

            "input",

            handlePromptInput

        );

    });

}

/* ============================================================
   INPUT EVENT
============================================================ */

function handlePromptInput(event) {

    AppState.dirty = true;

    const textarea = event.target;

    let section = "offer";

    if (textarea === UI.context)
        section = "workbook";

    if (textarea === UI.sheetPrompt)
        section = "sheet";

    updatePromptStatus(
        section,
        "editing"
    );

    debounceSave(section);

}

/* ============================================================
   DEBOUNCE SAVE
============================================================ */

function debounceSave(section) {

    clearTimeout(
        AutoSave.timers[section]
    );

    AutoSave.timers[section] =
        setTimeout(() => {

            regeneratePrompt();

            saveDraft();

            AppState.dirty = false;

            updatePromptStatus(
                section,
                "ready"
            );

        }, 500);

}

/* ============================================================
   EXPAND CURRENT SECTION
============================================================ */

function expandSection(textarea) {

    const collapse =
        textarea.closest(".collapse");

    if (!collapse)
        return;

    document
        .querySelectorAll(
            ".prompt-builder-card .prompt-section .collapse.show"
        )
        .forEach(openSection => {

            if (openSection === collapse) return;

            bootstrap.Collapse
                .getOrCreateInstance(
                    openSection,
                    { toggle: false }
                )
                .hide();

        });

    bootstrap
        .Collapse
        .getOrCreateInstance(
            collapse,
            { toggle: false }
        )
        .show();

    const icon =
        collapse
            .closest(".prompt-section")
            ?.querySelector(
                ".accordion-toggle i"
            );

    if (icon) {

        icon.className =
            "bi bi-chevron-down";

    }

}

/* ============================================================
   COLLAPSE ICONS
============================================================ */

function initializeAccordions() {

    document
        .querySelectorAll(".collapse")
        .forEach(section => {

            section.addEventListener(

                "shown.bs.collapse",

                () => {

                    document
                        .querySelectorAll(
                            ".prompt-section.active-section"
                        )
                        .forEach(card => {

                            if (
                                card !==
                                section.closest(".prompt-section")
                            ) {

                                card.classList.remove(
                                    "active-section"
                                );

                            }

                        });

                    section
                        .closest(".prompt-section")
                        ?.classList.add("active-section");

                    const icon =
                        section
                            .closest(".prompt-section")
                            ?.querySelector(
                                ".accordion-toggle i"
                            );

                    if (icon)
                        icon.className =
                            "bi bi-chevron-down";

                }

            );

            section.addEventListener(

                "hidden.bs.collapse",

                () => {

                    section
                        .closest(".prompt-section")
                        ?.classList.remove("active-section");

                    const icon =
                        section
                            .closest(".prompt-section")
                            ?.querySelector(
                                ".accordion-toggle i"
                            );

                    if (icon)
                        icon.className =
                            "bi bi-chevron-right";

                }

            );

        });

}

/* ============================================================
   RESET EDIT MODE
============================================================ */

function resetEditors() {

    EditState.offer = false;
    EditState.workbook = false;
    EditState.sheet = false;

    UI.offer.readOnly = true;
    UI.context.readOnly = true;
    UI.sheetPrompt.readOnly = true;

}

/* ============================================================
   KEYBOARD SHORTCUTS
============================================================ */

document.addEventListener(

    "keydown",

    event => {

        /* CTRL + S */

        if (

            event.ctrlKey &&

            event.key.toLowerCase() === "s"

        ) {

            event.preventDefault();

            regeneratePrompt();

            saveDraft();

            showToast(
                "Draft Saved"
            );

        }

        /* CTRL + ENTER */

        if (

            event.ctrlKey &&

            event.key === "Enter"

        ) {

            event.preventDefault();

            regeneratePrompt();

            showToast(
                "Prompt Updated"
            );

        }

    }

);

/* ============================================================
   INITIALIZE PROMPT BUILDER
============================================================ */

function initializePromptBuilder() {

    registerPromptBuilderEvents();

    registerPromptInputEvents();

    initializeAccordions();

    resetEditors();

}

/******************************************************************
 * 04_promptGeneration.js
 * Part 4A
 * Prompt Generation Engine
 ******************************************************************/

"use strict";

/* ============================================================
   REGENERATE PROMPT
============================================================ */

function hasCompletePromptData() {

    return Boolean(
        AppState.selectedSheets.length &&
        AppState.temporaryWorkbook &&
        UI.offer?.value.trim() &&
        UI.sheetPrompt?.value.trim()
    );

}

function updateGeneratedPromptStatus(ready) {

    if (UI.generatedPromptStatus) {

        UI.generatedPromptStatus.className =
            `compact-status text-${ready ? "success" : "secondary"}`;

        UI.generatedPromptStatus.textContent =
            "●";

        UI.generatedPromptStatus.title =
            ready ? "Ready" : "Waiting";

        UI.generatedPromptStatus.textContent =
            ready ? "● Ready" : "● Waiting";

        UI.generatedPromptStatus.setAttribute(
            "aria-label",
            ready ? "Ready" : "Waiting"
        );

    }

    updatePromptReadiness(
        ready ? "safety-pending" : "waiting"
    );

    updateGPTAvailability(ready);

    const safetyCard =
        document.querySelector(
            ".prompt-safety-card"
        );

    safetyCard?.classList.toggle(
        "safety-disabled",
        false
    );

    const safetyToggle =
        safetyCard?.querySelector(
            '[data-bs-target="#safetyCollapse"]'
        );

    if (safetyToggle) {

        safetyToggle.disabled = false;
        safetyToggle.setAttribute(
            "aria-disabled",
            "false"
        );

    }

    if (UI.previewButton) {

        UI.previewButton.disabled = !ready;

    }

    if (
        ready &&
        AppState.promptSafe &&
        UI.gptIframe &&
        UI.gptIframe.getAttribute("src") === "about:blank"
    ) {

        UI.gptIframe.src = "https://chatgpt.com";

    }

}

function updateGPTAvailability(promptReady) {

    const enabled =
        Boolean(promptReady && AppState.promptSafe);

    UI.gptWorkspaceEmpty?.classList.toggle(
        "d-none",
        enabled
    );

    UI.gptIframe?.classList.toggle(
        "d-none",
        !enabled
    );

    UI.copyButton?.classList.toggle(
        "d-none",
        !enabled
    );

    if (
        enabled &&
        UI.gptIframe?.getAttribute("src") === "about:blank"
    ) {

        UI.gptIframe.src = "https://chatgpt.com";

    }

}

function updatePromptReadiness(state) {

    if (!UI.promptReadiness) return;

    const states = {
        waiting: ["Load a workbook to build your prompt", "secondary"],
        "safety-pending": ["🟡 Safety Pending", "warning"],
        ready: ["🟢 Ready for GPT", "success"],
        pii: ["🔴 PII Detected", "danger"]
    };

    const [message, color] =
        states[state] || states.waiting;

    UI.promptReadiness.className =
        `prompt-readiness text-${color}`;

    UI.promptReadiness.textContent = message;

}

function regeneratePrompt() {

    const offer =
        UI.offer?.value.trim() || "";

    const context =
        UI.context?.value.trim() || "";

    const instructions =
        UI.sheetPrompt?.value.trim() || "";

    AppState.prompt.offer = offer;
    AppState.prompt.context = context;
    AppState.prompt.sheet = instructions;

    const ready = hasCompletePromptData();

    const unmaskedPrompt = ready
        ? buildPrompt(
            offer,
            instructions
        )
        : "";

    const maskingResult =
        applyAutomaticMasking(unmaskedPrompt);

    const prompt = maskingResult.text;
    AppState.maskingSummary = maskingResult.summary;
    AppState.prompt.final = prompt;

    if (UI.finalPrompt) {

        UI.finalPrompt.value = prompt;

    }

    syncPromptPreview();

    updatePromptMetrics();

    updateGeneratedPromptStatus(ready);

    resetPromptSafety();

}

function applyAutomaticMasking(text) {

    const summary = [];
    const patterns = [
        {
            type: "Credit Card",
            regex: /(?<!\d)(?:\d{16}|\d{4}-\d{4}-\d{4}-\d{4}|\d{4}\.\d{4}\.\d{4}\.\d{4}|\d{4}\s\d{4}\s\d{4}\s\d{4})(?!\d)/g,
            masked: "[MASKED CREDIT CARD]"
        },
        {
            type: "SSN",
            regex: /(?<!\d)(?:\d{3}-\d{2}-\d{4}|\d{3}\.\d{2}\.\d{4}|\d{3}\s\d{2}\s\d{4})(?!\d)/g,
            masked: "[MASKED SSN]"
        }
    ];

    let maskedText = text;

    patterns.forEach(({ type, regex, masked }) => {

        maskedText = maskedText.replace(regex, original => {

            summary.push({
                type,
                original,
                masked
            });

            return masked;

        });

    });

    return { text: maskedText, summary };

}

function renderMaskingSummary() {

    if (
        !UI.maskingSummary ||
        !UI.maskingSummaryCount ||
        !UI.maskingSummaryList
    ) return;

    UI.maskingSummaryCount.textContent =
        AppState.maskingSummary.length;
    UI.maskingSummaryList.replaceChildren();

    const creditCardCount =
        AppState.maskingSummary.filter(
            item => item.type === "Credit Card"
        ).length;
    const ssnCount =
        AppState.maskingSummary.filter(
            item => item.type === "SSN"
        ).length;

    if (UI.creditCardMaskStatus)
        UI.creditCardMaskStatus.textContent =
            `${creditCardCount} masked`;
    if (UI.ssnMaskStatus)
        UI.ssnMaskStatus.textContent =
            `${ssnCount} masked`;

    AppState.maskingSummary.forEach(item => {

        const row = document.createElement("div");
        row.className = "masking-summary-row";

        const type = document.createElement("span");
        type.className = "badge bg-secondary";
        type.textContent = item.type;

        const values = document.createElement("code");
        values.textContent =
            `${item.original} \u2192 ${item.masked}`;

        row.append(type, values);
        UI.maskingSummaryList.appendChild(row);

    });

    UI.maskingSummary.classList.remove("d-none");

}

/* ============================================================
   BUILD PROMPT
============================================================ */

function buildPrompt(

    offer,

    instructions

) {

    const attachment =
        AppState.temporaryWorkbook;

    return [

        "# Workbook Validation Prompt",

        "",

        "## Offer Description",

        offer || "Not Available",

        "",

        "## Validation Instructions",

        instructions || "Not Available",

        "",

        "## Attached Workbook",

        `Use the attached workbook "${attachment?.filename || "temporary workbook"}".`,

        `Included worksheets: ${attachment?.sheets?.join(", ") || "None"}.`,

        "Validate only the attached workbook. Worksheet data is intentionally not embedded in this prompt."

    ].join("\n");

}

/* ============================================================
   SYNC PREVIEW
============================================================ */

function syncPromptPreview() {

    if (

        UI.previewArea &&

        UI.finalPrompt

    ) {

        UI.previewArea.value =

            UI.finalPrompt.value;

    }

}

/* ============================================================
   PROMPT METRICS
============================================================ */

function updatePromptMetrics() {

    const prompt =

        UI.finalPrompt?.value || "";

    AppState.prompt.characters =

        prompt.length;

    AppState.prompt.words =

        prompt

            .trim()

            .split(/\s+/)

            .filter(Boolean)

            .length;

}

/* ============================================================
   AUTO REGENERATE
============================================================ */

function refreshGeneratedPrompt() {

    regeneratePrompt();

    saveDraft();

}

/* ============================================================
   CLEAR GENERATED PROMPT
============================================================ */

function clearGeneratedPrompt() {

    AppState.prompt.final = "";
    AppState.maskingSummary = [];

    if (UI.finalPrompt)

        UI.finalPrompt.value = "";

    if (UI.previewArea)

        UI.previewArea.value = "";

    updateGeneratedPromptStatus(false);

    resetPromptSafety();

}

/******************************************************************
 * Part 4B
 * Preview
 * Copy
 * Download
 * GPT Workspace
 ******************************************************************/

"use strict";

/* ============================================================
   SHOW PROMPT PREVIEW
============================================================ */

async function showPromptPreview() {

    regeneratePrompt();

    const safe = await validatePrompt();

    if (!safe) return;

    syncPromptPreview();

    BootstrapUI.previewModal?.show();

    setWorkflowStep(6);

    updateAIStatus("GPT Ready");

}

/* ============================================================
   CLOSE PREVIEW
============================================================ */

function closePromptPreview() {

    BootstrapUI.previewModal?.hide();

}

/* ============================================================
   COPY GENERATED PROMPT
============================================================ */

async function copyPrompt() {

    regeneratePrompt();

    const safe = await validatePrompt();

    if (!safe) return;

    await copyToClipboard(

        UI.finalPrompt.value,
        "✓ Generated prompt copied successfully."

    );

    updateAIStatus("Prompt Copied");

}

/* ============================================================
   COPY PREVIEW
============================================================ */

async function copyPreviewPrompt() {

    await copyToClipboard(

        UI.previewArea.value,
        "✓ Generated prompt copied successfully."

    );

}

/* ============================================================
   DOWNLOAD PROMPT
============================================================ */

function downloadPrompt() {

    regeneratePrompt();

    const blob = new Blob(

        [UI.finalPrompt.value],

        {

            type: "text/plain"

        }

    );

    const url =

        URL.createObjectURL(blob);

    const link =

        document.createElement("a");

    link.href = url;

    link.download =

        "Workbook_AI_Prompt.txt";

    link.click();

    URL.revokeObjectURL(url);

}

/* ============================================================
   OPEN GPT
============================================================ */

function openGPTWorkspace() {

    const iframe =

        document.getElementById(

            "gptIframe"

        );

    if (iframe) {

        iframe.src =

            "https://chatgpt.com";

        return;

    }

    window.open(

        "https://chatgpt.com",

        "_blank"

    );

}

/* ============================================================
   VALIDATION
============================================================ */

async function validatePrompt() {

    if (

        !UI.finalPrompt ||

        !UI.finalPrompt.value.trim()

    ) {

        showToast(

            "Generate a prompt first."

        );

        return false;

    }

    if (

        typeof scanPromptForPII !==

        "function"

    ) {

        return true;

    }

    return await scanPromptForPII();

}

/* ============================================================
   PROMPT SUMMARY
============================================================ */

function promptSummary() {

    return {

        characters:

            AppState.prompt.characters,

        words:

            AppState.prompt.words,

        sheet:

            AppState.selectedSheet,

        workbook:

            AppState.workbookName

    };

}

/* ============================================================
   BUTTON EVENTS
============================================================ */

function registerPromptActions() {

    UI.previewButton?.addEventListener(

        "click",

        showPromptPreview

    );

    UI.copyButton?.addEventListener(

        "click",

        copyPrompt

    );

    UI.copyPreviewButton?.addEventListener(

        "click",

        copyPreviewPrompt

    );

}

/* ============================================================
   INITIALIZE
============================================================ */

function initializePromptGeneration() {

    const actionFooter =
        document.querySelector(
            ".generated-prompt-card > .card-footer"
        );

    const headerRow =
        document.querySelector(
            ".generated-prompt-card > .card-header > .d-flex"
        );

    const actionBar =
        actionFooter?.querySelector(
            ".prompt-action-bar"
        );

    if (actionFooter && actionBar && headerRow) {

        const headerActions =
            document.createElement("div");

        headerActions.className =
            "generated-prompt-header-actions";

        const titleGroup =
            headerRow.firstElementChild;

        if (UI.generatedPromptStatus && titleGroup) {

            titleGroup.classList.add(
                "generated-prompt-title-group"
            );

            titleGroup.appendChild(
                UI.generatedPromptStatus
            );

        }

        actionBar.className =
            "generated-prompt-actions btn-group";

        UI.previewButton.innerHTML =
            '<i class="bi bi-eye me-1"></i>Preview Prompt';
        UI.previewButton.title = "Preview Prompt";
        UI.previewButton.setAttribute(
            "aria-label",
            "Preview Prompt"
        );

        UI.copyButton.innerHTML =
            '<i class="bi bi-clipboard me-1"></i>Copy Generated Prompt';
        UI.copyButton.title = "Copy Generated Prompt";
        UI.copyButton.setAttribute(
            "aria-label",
            "Copy Generated Prompt"
        );

        headerActions.appendChild(actionBar);
        headerRow.appendChild(headerActions);
        actionFooter.remove();

    }

    registerPromptActions();

    syncPromptPreview();

    updateGeneratedPromptStatus(
        hasCompletePromptData()
    );

    resetPromptSafety();

}

/******************************************************************
 * 05_piiWorkflow.js
 * Part 5A
 * PII Scanner
 * Prompt Validation
 ******************************************************************/

"use strict";

/* ============================================================
   PII SCAN
============================================================ */

function resetPromptSafety() {

    AppState.promptSafe = false;

    updateGPTAvailability(
        hasCompletePromptData()
    );

    setBadge(
        UI.promptSafetyStatus,
        "Active",
        "success"
    );

    if (UI.promptSafetyAlert) {

        UI.promptSafetyAlert.className =
            "alert alert-secondary mt-3 mb-0";

        UI.promptSafetyAlert.innerHTML =
            '<i class="bi bi-shield-exclamation me-2"></i>' +
            "<strong>Prompt Safety Notice:</strong> This tool automatically masks Credit Card and SSN values only. It does not detect or mask other sensitive or confidential information. Please review the generated prompt carefully before copying or sharing it externally.";

    }

    renderMaskingSummary();

}

function getPromptSections() {

    const sections = {};
    const selectedSheets =
        AppState.selectedSheets.length
            ? AppState.selectedSheets
            : [AppState.selectedSheet].filter(Boolean);

    if (selectedSheets.length) {

        sections[
            `${selectedSheets[0]}::Offer Description`
        ] = UI.offer?.value || "";

    }

    selectedSheets.forEach(name => {

        const sheet =
            AppState.workbook?.sheet_data?.[name];

        if (!sheet) return;

        sections[
            `${name}::Workbook Context`
        ] = sheet.workbook_context || "";

        sections[
            `${name}::Sheet Instructions`
        ] = [
            sheet.sheet_prompt || "",
            sheet.sheet_sample || ""
        ].join("\n\n");

    });

    return sections;

}

async function scanPromptForPII() {

    try {

        showLoading(false);

        setBadge(
            UI.promptSafetyStatus,
            "Validating",
            "warning",
            true
        );

        updateAIStatus("Scanning Prompt...");

        const response = await fetch(

            "/scan_pii",

            {

                method: "POST",

                headers: {

                    "Content-Type":
                        "application/json"

                },

                body: JSON.stringify({

                    prompt:

                        UI.finalPrompt.value

                })

            }

        );

        const result =
            await response.json();

        hideLoading();

        if (!response.ok) {

            throw new Error(

                result.error ||

                "PII Scan Failed"

            );

        }

        updatePIIInformation(result);

        return !result.has_pii;

    }

    catch (error) {

        hideLoading();

        console.error(error);

        showError(

            error.message

        );

        return false;

    }

}

/* ============================================================
   UPDATE PII STATUS
============================================================ */

function updatePIIInformation(result) {

    AppState.promptSafe = !result.has_pii;
    setBadge(
        UI.promptSafetyStatus,
        result.has_pii ? "Masking Required" : "Protected",
        result.has_pii ? "danger" : "success"
    );
    updatePromptReadiness(
        result.has_pii ? "pii" : "ready"
    );
    updateGPTAvailability(
        !result.has_pii && hasCompletePromptData()
    );
    renderMaskingSummary();
    updateAIStatus(
        result.has_pii ? "Masking Required" : "Prompt Safe"
    );

    return;

    if (result.has_pii) {

        UI.piiResult.className =

            "badge bg-danger";

        UI.piiResult.innerHTML =

            `${result.count} PII Found`;

    }

    else {

        UI.piiResult.className =

            "badge bg-success";

        UI.piiResult.innerHTML =

            "No PII Found";

    }

    setBadge(
        UI.promptSafetyStatus,
        result.has_pii
            ? `⚠ PII Detected (${result.count || 0})`
            : "🟢 Safe to Send",
        result.has_pii ? "danger" : "success"
    );

    updatePromptReadiness(
        result.has_pii ? "pii" : "ready"
    );

    updateGPTAvailability(
        hasCompletePromptData()
    );

    if (UI.validationResult) {

        UI.validationResult.className =

            result.has_pii

                ? "badge bg-danger"

                : "badge bg-success";

        UI.validationResult.innerHTML =

            result.has_pii

                ? "Failed"

                : "Passed";

    }

    if (UI.promptSafetyAlert) {

        UI.promptSafetyAlert.className =
            result.has_pii
                ? "alert alert-danger mt-3 mb-0"
                : "alert alert-success mt-3 mb-0";

        UI.promptSafetyAlert.innerHTML =
            result.has_pii
                ? '<i class="bi bi-exclamation-triangle-fill me-2"></i>PII was detected. Remove it before using the prompt.'
                : '<i class="bi bi-check-circle-fill me-2"></i>Only Credit Card and SSN are automatically masked. Review the entire prompt before sending it to ChatGPT.';

    }

    renderPIIFindings();

    if (result.has_pii) {

        const safetyElement =
            document.getElementById("safetyCollapse");

        if (safetyElement && window.bootstrap) {

            bootstrap.Collapse
                .getOrCreateInstance(
                    safetyElement,
                    { toggle: false }
                )
                .show();

        }

        document.querySelector(".prompt-safety-card")
            ?.scrollIntoView({
                behavior: "smooth",
                block: "nearest"
            });

    }

    updateAIStatus(

        result.has_pii

            ? "PII Detected"

            : "Prompt Safe"

    );

}

function renderPIIFindings() {

    if (!UI.piiFindings || !UI.piiReport) return;

    UI.piiFindings.replaceChildren();

    if (!AppState.piiFindings.length) {

        UI.piiReport.classList.add("d-none");
        return;

    }

    AppState.piiFindings.forEach((finding, index) => {

        const row = document.createElement("div");
        row.className = "pii-finding";

        const typeLabel = String(finding.type || "unknown")
            .replaceAll("_", " ")
            .replace(/\b\w/g, character => character.toUpperCase());

        row.innerHTML = `
            <div class="pii-finding-grid">
                <div><small class="text-muted d-block">PII Type</small><strong>${escapeHTML(typeLabel)}</strong></div>
                <div><small class="text-muted d-block">Worksheet</small>${escapeHTML(finding.worksheet || "-")}</div>
                <div><small class="text-muted d-block">Section</small>${escapeHTML(finding.section || "-")}</div>
                <div><small class="text-muted d-block">Masked Preview</small><code>${escapeHTML(finding.masked_preview || "****")}</code></div>
                <div><small class="text-muted d-block">Severity</small><span class="badge bg-${finding.severity === "High" ? "danger" : finding.severity === "Medium" ? "warning text-dark" : "secondary"}">${escapeHTML(finding.severity || "Low")}</span></div>
                <div class="d-flex gap-1">
                    <button class="btn btn-sm btn-outline-primary" data-pii-action="goto" data-index="${index}">Go To</button>
                    <button class="btn btn-sm btn-primary" data-pii-action="mask" data-index="${index}">Mask Selected</button>
                </div>
            </div>
        `;

        UI.piiFindings.appendChild(row);

    });

    UI.piiReport.classList.remove("d-none");

}

function escapeHTML(value) {

    const element = document.createElement("span");
    element.textContent = String(value);
    return element.innerHTML;

}

function sectionTarget(section) {

    const targets = {
        "Offer Description": {
            field: UI.offer,
            collapse: "offerCollapse"
        },
        "Workbook Context": {
            field: UI.context,
            collapse: "workbookCollapse"
        },
        "Sheet Instructions": {
            field: UI.sheetPrompt,
            collapse: "sheetCollapse"
        }
    };

    return targets[section];

}

function goToPIIFinding(finding) {

    if (
        finding.worksheet &&
        UI.sheetSelector &&
        UI.sheetSelector.value !== finding.worksheet
    ) {

        UI.sheetSelector.value = finding.worksheet;
        handleSheetSelection();

    }

    const target = sectionTarget(finding.section);
    if (!target?.field) return;

    const collapseElement =
        document.getElementById(target.collapse);

    if (collapseElement && window.bootstrap) {

        bootstrap.Collapse
            .getOrCreateInstance(
                collapseElement,
                { toggle: false }
            )
            .show();

    }

    target.field.classList.add("pii-highlight");
    target.field.focus();
    target.field.setSelectionRange(
        finding.start,
        finding.end
    );
    target.field.scrollIntoView({
        behavior: "smooth",
        block: "center"
    });

    window.setTimeout(
        () => target.field.classList.remove("pii-highlight"),
        2400
    );

}

function maskFindings(findings) {

    const bySource = {};

    findings.forEach(finding => {

        const key =
            `${finding.worksheet}::${finding.section}`;

        (bySource[key] ||= []).push(finding);

    });

    Object.entries(bySource).forEach(
        ([key, sourceFindings]) => {

            const [worksheet, section] =
                key.split("::");

            const sheet =
                AppState.workbook
                    ?.sheet_data?.[worksheet];

            if (!sheet) return;

            let value =
                section === "Offer Description"
                    ? sheet.offer_description || ""
                    : section === "Workbook Context"
                        ? sheet.workbook_context || ""
                        : [
                            sheet.sheet_prompt || "",
                            sheet.sheet_sample || ""
                        ].join("\n\n");

            sourceFindings
                .sort((left, right) => right.start - left.start)
                .forEach(finding => {

                    value =
                        value.slice(0, finding.start) +
                        "*".repeat(finding.end - finding.start) +
                        value.slice(finding.end);

                });

            if (section === "Offer Description") {

                Object.values(
                    AppState.workbook.sheet_data
                ).forEach(item => {

                    item.offer_description = value;

                });

            }

            else if (section === "Workbook Context") {

                sheet.workbook_context = value;

            }

            else {

                const separator = "\n\n";
                const originalPrompt =
                    sheet.sheet_prompt || "";
                const splitAt =
                    Math.min(
                        originalPrompt.length,
                        value.length
                    );

                sheet.sheet_prompt =
                    value.slice(0, splitAt);
                sheet.sheet_sample =
                    value.slice(
                        splitAt + separator.length
                    );

            }

        }
    );

    loadSelectedSheets(
        AppState.selectedSheets
    );

    scanPromptForPII();

    return;

    const grouped = {};

    findings.forEach(finding => {

        (grouped[finding.section] ||= [])
            .push(finding);

    });

    Object.entries(grouped).forEach(
        ([section, sectionFindings]) => {

            const field =
                sectionTarget(section)?.field;

            if (!field) return;

            let value = field.value;

            sectionFindings
                .sort((left, right) => right.start - left.start)
                .forEach(finding => {

                    value =
                        value.slice(0, finding.start) +
                        "*".repeat(finding.end - finding.start) +
                        value.slice(finding.end);

                });

            field.value = value;

        }
    );

    regeneratePrompt();
    scanPromptForPII();

}

function registerPIIActions() {

    UI.piiFindings?.addEventListener("click", event => {

        const button =
            event.target.closest("[data-pii-action]");

        if (!button) return;

        const finding =
            AppState.piiFindings[
                Number(button.dataset.index)
            ];

        if (!finding) return;

        if (button.dataset.piiAction === "goto") {

            goToPIIFinding(finding);

        }

        if (button.dataset.piiAction === "mask") {

            maskFindings([finding]);

        }

    });

    UI.maskAllPiiButton?.addEventListener(
        "click",
        () => maskFindings([...AppState.piiFindings])
    );

    UI.ignoreAllPiiButton?.addEventListener(
        "click",
        () => {

            AppState.ignorePII = true;
            AppState.promptSafe = true;

            setBadge(
                UI.promptSafetyStatus,
                "PII Ignored",
                "warning"
            );

            UI.piiReport?.classList.add("d-none");

            updatePromptReadiness("ready");

            updateGPTAvailability(
                hasCompletePromptData()
            );

            updateAIStatus("PII Ignored");

        }
    );

}

/* ============================================================
   PROMPT SAFETY
============================================================ */

function isPromptSafe(result) {

    return !result.has_pii;

}

/* ============================================================
   VALIDATION SUMMARY
============================================================ */

function validationSummary(result) {

    return {

        safe:

            !result.has_pii,

        piiCount:

            result.count || 0,

        types:

            result.types || {}

    };

}

/******************************************************************
 * Part 5B
 * Error Handling
 * Application Reset
 * Initialization
 ******************************************************************/

"use strict";

/* ============================================================
   ERROR HANDLER
============================================================ */

function showError(message) {

    console.error(message);

    updateAIStatus("Error");

    showToast(message || "Unexpected error occurred.");

}

/* ============================================================
   RESET APPLICATION
============================================================ */

function resetApplication() {

    cancelWorkbookMonitoring();
    endWorkbookSession();
    BootstrapUI.sheetPreviewModal?.hide();
    BootstrapUI.previewModal?.hide();

    clearWorkbook();

    clearGeneratedPrompt();

    clearDraft();

    AppState.dirty = false;
    AppState.maskingSummary = [];
    AppState.promptSafe = false;
    AppState.temporaryWorkbook = null;

    AppState.selectedSheet = "";

    AppState.selectedSheets = [];
    AppState.worksheetSelectionOrder = [];

    AppState.prompt = {

        offer: "",

        context: "",

        sheet: "",

        final: "",

        characters: 0,

        words: 0

    };

    resetEditors();

    if (UI.projectRef) UI.projectRef.value = "";
    if (UI.taskNumber) UI.taskNumber.value = "";
    if (UI.manualWorkbookInput)
        UI.manualWorkbookInput.value = "";
    UI.manualDownloadPanel?.classList.add("d-none");

    if (UI.previewSheetButton)
        UI.previewSheetButton.disabled = true;
    renderMaskingSummary();
    clearTemporaryWorkbookAttachment();
    resetPromptSafety();

    const safetyCollapse =
        document.getElementById("safetyCollapse");
    if (safetyCollapse && window.bootstrap) {

        bootstrap.Collapse
            .getOrCreateInstance(
                safetyCollapse,
                { toggle: false }
            )
            .show();

    }

    updateAIStatus("Ready");

    updatePIIBadge(false);

    showToast("Workspace reset successfully.");

}

function endWorkbookSession() {

    const sessionId =
        AppState.workbook?.workbook_session_id;

    if (!sessionId) return;

    fetch(
        "/workbook_session/end",
        {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                workbook_session_id: sessionId
            })
        }
    ).catch(() => {});

}

/* ============================================================
   ENABLE INPUTS
============================================================ */

function enablePromptBuilder() {

    [

        UI.offer,

        UI.context,

        UI.sheetPrompt,

        UI.finalPrompt

    ].forEach(item => {

        if (item)

            item.disabled = false;

    });

}

/* ============================================================
   DISABLE INPUTS
============================================================ */

function disablePromptBuilder() {

    [

        UI.offer,

        UI.context,

        UI.sheetPrompt,

        UI.finalPrompt

    ].forEach(item => {

        if (item)

            item.disabled = true;

    });

}

/* ============================================================
   GLOBAL ERROR EVENTS
============================================================ */

window.addEventListener(

    "error",

    event => {

        console.error(event.error);

        showError(

            event.message

        );

    }

);

window.addEventListener(

    "unhandledrejection",

    event => {

        console.error(

            event.reason

        );

        showError(

            "Unexpected application error."

        );

    }

);

window.addEventListener("pagehide", () => {

    const sessionId =
        AppState.workbook?.workbook_session_id;
    if (!sessionId) return;

    fetch(
        "/workbook_session/end",
        {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                workbook_session_id: sessionId
            }),
            keepalive: true
        }
    ).catch(() => {});

});

/* ============================================================
   APPLICATION STARTUP
============================================================ */

document.addEventListener(

    "DOMContentLoaded",

    () => {

        initializeApplication();

        initializePromptBuilder();

        initializePromptGeneration();

        disablePromptBuilder();

        updateAIStatus("Ready");

        console.log(

            "Workbook Validation Tool Ready"

        );

    }

);

/* ============================================================
   PUBLIC API
============================================================ */

window.WorkbookApp = {

    fetchWorkbook,

    regeneratePrompt,

    showPromptPreview,

    copyPrompt,

    copyPreviewPrompt,

    refreshCurrentSheet,

    resetApplication,

    scanPromptForPII

};
