const el = (id) => document.getElementById(id);
const appBase = (document.querySelector('meta[name="app-base"]')?.content || "").replace(/\/$/, "");
const appUrl = (path) => `${appBase}${path.startsWith("/") ? path : `/${path}`}`;
const ui = {
    scanView: el("scanView"), processingView: el("processingView"), resultsView: el("resultsView"),
    drop: el("dropZone"), input: el("pdfFile"), choose: el("chooseBtn"),
    meta: el("fileMeta"), fileName: el("fileName"), fileSize: el("fileSize"), fileId: el("currentFileId"), scanError: el("scanError"),
    processingPanel: el("processingPanel"), scanner: el("qrScanner"), processingTitle: el("processingTitle"), processingError: el("processingError"), processingErrorText: el("processingErrorText"), errorReturn: el("errorReturnBtn"),
    fill: el("progressFill"), pct: el("progressPct"), status: el("progressStatus"), pages: el("progressPages"), found: el("liveQr"), processed: el("pagesProcessed"), elapsed: el("elapsedTime"),
    total: el("totalPages"), summaryFound: el("qrFound"), summaryTime: el("summaryTime"), count: el("resultsCount"),
    search: el("resultSearch"), sort: el("resultSort"), grid: el("qrGrid"), empty: el("qrEmpty"), emptyTitle: el("emptyTitle"), emptyText: el("emptyText"), pagination: el("pagination"),
    scanAnother: el("scanAnotherBtn"), emptyScanAnother: el("emptyScanAnotherBtn"), theme: el("themeToggle"),
    modal: el("qrModal"), modalClose: el("modalClose"), modalImage: el("modalImage"), modalPage: el("modalPage"), modalType: el("modalType"), modalPayload: el("modalPayload"), modalOpenLink: el("modalOpenLink")
};

let selectedFile = null;
let items = [];
let pageIndex = 1;
let progressTimer = null;
let activeController = null;
let modalTrigger = null;
const pageSize = 12;

function showView(name) {
    ui.scanView.hidden = name !== "scan";
    ui.processingView.hidden = name !== "processing";
    ui.resultsView.hidden = name !== "results";
    window.scrollTo({ top: 0, behavior: "smooth" });
}

function formatBytes(bytes) { return `${(bytes / 1024 / 1024).toFixed(2)} MB`; }
function formatTime(seconds) { const value = Math.max(0, Math.round(Number(seconds) || 0)); return `${String(Math.floor(value / 60)).padStart(2, "0")}:${String(value % 60).padStart(2, "0")}`; }
function typeLabel(item) { return item.payload_type || "Text"; }
function readJson(response) { return response.json().catch(() => ({})); }
function wait(milliseconds) { return new Promise((resolve) => window.setTimeout(resolve, milliseconds)); }
function webUrlFromPayload(payload) {
    const value = String(payload || "");
    if (!value || value !== value.trim() || /\s/.test(value)) return null;
    const hasWebProtocol = /^https?:\/\//i.test(value);
    if (!hasWebProtocol && value.includes("@")) return null;
    const destination = hasWebProtocol ? value : `https://${value}`;
    try {
        const parsed = new URL(destination);
        if (parsed.protocol !== "http:" && parsed.protocol !== "https:") return null;
        const labels = parsed.hostname.split(".");
        if (labels.length < 2 || labels.some((label) => !/^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?$/i.test(label))) return null;
        const topLevelDomain = labels[labels.length - 1];
        if (!/^[a-z]{2,63}$/i.test(topLevelDomain) && !/^xn--[a-z0-9-]{2,59}$/i.test(topLevelDomain)) return null;
        return destination;
    } catch (_error) {
        return null;
    }
}
function showScanError(message) { ui.scanError.textContent = message; ui.scanError.hidden = false; }
function clearScanError() { ui.scanError.textContent = ""; ui.scanError.hidden = true; }

function resetProgress() {
    ui.fill.style.width = "0%";
    ui.fill.parentElement.setAttribute("aria-valuenow", "0");
    ui.pct.textContent = "0%";
    ui.status.textContent = "Preparing your document for QR code detection";
    ui.pages.textContent = "0 of 0";
    ui.processed.textContent = "0";
    ui.found.textContent = "0";
    ui.elapsed.textContent = "00:00";
    ui.processingTitle.textContent = "Scanning PDF...";
    ui.processingPanel.classList.remove("is-success", "is-error");
    ui.processingError.hidden = true;
}

function resetApplication() {
    if (activeController) activeController.abort();
    if (progressTimer) window.clearTimeout(progressTimer);
    activeController = null;
    progressTimer = null;
    selectedFile = null;
    items = [];
    pageIndex = 1;
    ui.input.value = "";
    ui.fileId.value = "";
    ui.fileName.textContent = "";
    ui.fileSize.textContent = "";
    ui.meta.hidden = true;
    ui.search.value = "";
    ui.sort.value = "asc";
    ui.grid.replaceChildren();
    ui.pagination.replaceChildren();
    ui.empty.hidden = true;
    clearScanError();
    resetProgress();
    closeModal(false);
    showView("scan");
    ui.drop.focus();
}

function chooseFile(file) {
    clearScanError();
    if (!file) return;
    if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) { resetApplication(); return showScanError("Only PDF files are supported."); }
    if (file.size > 30 * 1024 * 1024) { resetApplication(); return showScanError("The PDF must be 30 MB or smaller."); }
    selectedFile = file;
    ui.fileName.textContent = file.name;
    ui.fileSize.textContent = formatBytes(file.size);
    ui.meta.hidden = false;
    scanPdf();
}

function updateProgress(state) {
    const total = Number(state.total_pages) || 0;
    const processed = Number(state.pages_processed) || 0;
    const current = total ? Math.min(Number(state.current_page) || 0, total) : 0;
    const percent = total ? Math.min(100, Math.round(processed / total * 100)) : 0;
    ui.fill.style.width = `${percent}%`;
    ui.fill.parentElement.setAttribute("aria-valuenow", String(percent));
    ui.pct.textContent = `${percent}%`;
    ui.status.textContent = total ? "Detecting QR codes across your PDF" : "Preparing your document for QR code detection";
    ui.pages.textContent = `${current} of ${total}`;
    ui.found.textContent = String(Number(state.qr_count) || 0);
    ui.processed.textContent = String(processed);
    ui.elapsed.textContent = formatTime(state.elapsed_seconds);
}

function waitForNextPoll(signal) {
    return new Promise((resolve, reject) => {
        const onAbort = () => { window.clearTimeout(progressTimer); reject(new DOMException("Aborted", "AbortError")); };
        signal.addEventListener("abort", onAbort, { once: true });
        progressTimer = window.setTimeout(() => {
            signal.removeEventListener("abort", onAbort);
            progressTimer = null;
            resolve();
        }, 400);
    });
}

async function pollProgress(fileId, signal) {
    while (!signal.aborted) {
        const response = await fetch(appUrl(`/api/progress/${fileId}`), { cache: "no-store", signal });
        const state = await readJson(response);
        if (!response.ok) throw new Error("Unable to update scan progress.");
        updateProgress(state);
        const status = String(state.status || (state.done ? "COMPLETED" : "PROCESSING")).toUpperCase();
        if (status === "FAILED") throw new Error(state.error || "The PDF scan failed.");
        if (status === "COMPLETED") return state;
        await waitForNextPoll(signal);
    }
    throw new DOMException("Aborted", "AbortError");
}

function showProcessingError(message) {
    ui.processingPanel.classList.remove("is-success");
    ui.processingPanel.classList.add("is-error");
    ui.processingTitle.textContent = "Scan unsuccessful";
    ui.status.textContent = "The PDF could not be processed.";
    ui.processingErrorText.textContent = message || "Please return and try another PDF.";
    ui.processingError.hidden = false;
}

async function scanPdf() {
    if (!selectedFile) return;
    let uploadCompleted = false;
    clearScanError();
    resetProgress();
    showView("processing");
    activeController = new AbortController();
    const { signal } = activeController;
    try {
        const form = new FormData();
        form.append("file", selectedFile);
        const uploadResponse = await fetch(appUrl("/api/upload"), { method: "POST", body: form, signal });
        const upload = await readJson(uploadResponse);
        if (!uploadResponse.ok) throw new Error(upload.error || "Unable to upload the PDF.");
        uploadCompleted = true;
        ui.fileId.value = upload.file_id;
        const startResponse = await fetch(appUrl(`/api/process/${upload.file_id}`), { signal });
        const started = await readJson(startResponse);
        if (!startResponse.ok) throw new Error(started.error || "Unable to start PDF processing.");
        let result = started;
        if (startResponse.status === 202) {
            await pollProgress(upload.file_id, signal);
            const resultResponse = await fetch(appUrl(`/api/process/${upload.file_id}`), { cache: "no-store", signal });
            result = await readJson(resultResponse);
            if (!resultResponse.ok) throw new Error(result.error || "Unable to retrieve scan results.");
        }
        updateProgress({ ...result, current_page: result.total_pages, pages_processed: result.total_pages, elapsed_seconds: result.scan_time_seconds, done: true });
        items = result.qr_items || [];
        ui.total.textContent = String(result.total_pages || 0);
        ui.summaryFound.textContent = String(result.qr_count || 0);
        ui.summaryTime.textContent = formatTime(result.scan_time_seconds);
        ui.processingPanel.classList.add("is-success");
        ui.processingPanel.classList.remove("is-error");
        ui.processingError.hidden = true;
        ui.processingTitle.textContent = "Scan complete";
        ui.status.textContent = "QR code detection finished successfully";
        await wait(650);
        pageIndex = 1;
        renderResults();
        showView("results");
    } catch (error) {
        if (error.name !== "AbortError") {
            if (!uploadCompleted) {
                const message = error.message || "Unable to upload the PDF.";
                resetApplication();
                showScanError(message);
            } else {
                showProcessingError(error.message || "Processing failed. Please try again.");
            }
        }
    } finally {
        if (progressTimer) window.clearTimeout(progressTimer);
        progressTimer = null;
        if (activeController) activeController.abort();
        activeController = null;
    }
}

function filteredItems() {
    const query = ui.search.value.trim().toLowerCase();
    const filtered = items.filter((item) => String(item.payload || "").toLowerCase().includes(query));
    filtered.sort((a, b) => (Number(a.page) - Number(b.page)) * (ui.sort.value === "desc" ? -1 : 1));
    return filtered;
}

function renderResults() {
    const filtered = filteredItems();
    const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
    pageIndex = Math.min(pageIndex, totalPages);
    const visible = filtered.slice((pageIndex - 1) * pageSize, pageIndex * pageSize);
    ui.count.textContent = String(filtered.length);
    ui.grid.replaceChildren();
    const hasQuery = Boolean(ui.search.value.trim());
    ui.empty.hidden = filtered.length !== 0;
    ui.emptyTitle.textContent = hasQuery ? "No Matching QR Codes" : "No QR Codes Found";
    ui.emptyText.textContent = hasQuery ? "No QR payloads match your search." : "No QR codes were detected in this PDF.";
    ui.emptyScanAnother.hidden = hasQuery;
    visible.forEach((item) => {
        const card = document.createElement("article");
        card.className = "qr-card";
        const image = document.createElement("img");
        image.className = "qr-img";
        image.src = item.preview ? appUrl(item.preview) : "";
        image.alt = `Open QR code details from page ${item.page}`;
        image.tabIndex = 0;
        image.setAttribute("role", "button");
        image.addEventListener("click", () => openModal(item, image));
        image.addEventListener("keydown", (event) => { if (event.key === "Enter" || event.key === " ") { event.preventDefault(); openModal(item, image); } });
        image.addEventListener("error", () => { image.classList.add("preview-error"); image.alt = `QR preview unavailable for page ${item.page}`; image.removeAttribute("role"); image.removeAttribute("tabindex"); });
        const details = document.createElement("div");
        details.className = "qr-card-details";
        details.innerHTML = `<span class="page-badge">Page ${Number(item.page) || 0}</span><h3>QR Code</h3><span class="detail-label">Payload Type</span>`;
        const type = document.createElement("strong"); type.textContent = typeLabel(item);
        const label = document.createElement("span"); label.className = "detail-label"; label.textContent = "Payload Details";
        const payload = document.createElement("p"); payload.textContent = item.payload || "";
        details.append(type, label, payload);
        const linkUrl = webUrlFromPayload(item.payload);
        if (linkUrl) {
            const openLink = document.createElement("a");
            openLink.className = "open-link card-open-link";
            openLink.textContent = "Open Link ↗";
            openLink.setAttribute("href", linkUrl);
            openLink.setAttribute("target", "_blank");
            openLink.setAttribute("rel", "noopener noreferrer");
            details.appendChild(openLink);
        }
        card.append(image, details);
        ui.grid.appendChild(card);
    });
    renderPagination(totalPages, filtered.length);
}

function renderPagination(totalPages, resultCount) {
    ui.pagination.replaceChildren();
    if (totalPages <= 1 || resultCount === 0) return;
    const addButton = (label, number, className = "") => {
        const button = document.createElement("button"); button.type = "button"; button.textContent = label; button.className = className; button.disabled = number < 1 || number > totalPages;
        button.addEventListener("click", () => { pageIndex = number; renderResults(); document.querySelector(".results-head").scrollIntoView({ behavior: "smooth" }); });
        ui.pagination.appendChild(button);
    };
    addButton("‹", pageIndex - 1);
    for (let number = 1; number <= totalPages; number += 1) addButton(String(number), number, number === pageIndex ? "active" : "");
    addButton("›", pageIndex + 1);
}

function openModal(item, trigger) {
    modalTrigger = trigger || null;
    ui.modalImage.src = item.preview ? appUrl(item.preview) : "";
    ui.modalImage.alt = `QR code from page ${item.page}`;
    ui.modalPage.textContent = String(item.page);
    ui.modalType.textContent = typeLabel(item);
    ui.modalPayload.textContent = item.payload || "";
    const linkUrl = webUrlFromPayload(item.payload);
    ui.modalOpenLink.hidden = !linkUrl;
    if (linkUrl) ui.modalOpenLink.setAttribute("href", linkUrl);
    else ui.modalOpenLink.removeAttribute("href");
    ui.modal.hidden = false;
    document.body.classList.add("modal-open");
    ui.modalClose.focus();
}

function closeModal(restoreFocus = true) {
    if (ui.modal.hidden) return;
    ui.modal.classList.add("is-closing");
    window.setTimeout(() => { ui.modal.hidden = true; ui.modal.classList.remove("is-closing"); }, 150);
    document.body.classList.remove("modal-open");
    if (restoreFocus && modalTrigger) modalTrigger.focus();
    modalTrigger = null;
}

ui.choose.addEventListener("click", (event) => { event.stopPropagation(); ui.input.click(); });
ui.drop.addEventListener("click", () => ui.input.click());
ui.drop.addEventListener("keydown", (event) => { if (event.key === "Enter" || event.key === " ") { event.preventDefault(); ui.input.click(); } });
ui.input.addEventListener("change", () => chooseFile(ui.input.files[0]));
ui.drop.addEventListener("dragover", (event) => { event.preventDefault(); ui.drop.classList.add("dragover"); });
ui.drop.addEventListener("dragleave", () => ui.drop.classList.remove("dragover"));
ui.drop.addEventListener("drop", (event) => { event.preventDefault(); ui.drop.classList.remove("dragover"); chooseFile(event.dataTransfer.files[0]); });
ui.errorReturn.addEventListener("click", resetApplication);
ui.scanAnother.addEventListener("click", resetApplication);
ui.emptyScanAnother.addEventListener("click", resetApplication);
ui.search.addEventListener("input", () => { pageIndex = 1; renderResults(); });
ui.sort.addEventListener("change", () => { pageIndex = 1; renderResults(); });
ui.modalClose.addEventListener("click", () => closeModal());
ui.modal.addEventListener("click", (event) => { if (event.target === ui.modal) closeModal(); });
document.addEventListener("keydown", (event) => { if (event.key === "Escape" && !ui.modal.hidden) closeModal(); });

function setTheme(dark) { document.body.classList.toggle("dark-mode", dark); ui.theme.textContent = dark ? "☀" : "☾"; ui.theme.setAttribute("aria-label", dark ? "Use light mode" : "Use dark mode"); }
setTheme(localStorage.getItem("qr-theme") === "dark");
ui.theme.addEventListener("click", () => { const dark = !document.body.classList.contains("dark-mode"); setTheme(dark); localStorage.setItem("qr-theme", dark ? "dark" : "light"); });
showView("scan");
