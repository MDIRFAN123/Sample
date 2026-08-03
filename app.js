const el = (id) => document.getElementById(id);
const RING_CIRCUMFERENCE = 2 * Math.PI * 94;

// Enhanced: cache the expanded UI once; API and legacy IDs remain unchanged.
const ui = {
    uploadSection: el("uploadSection"), results: el("results"), kpis: el("kpiGrid"),
    successBanner: el("successBanner"), successSummary: el("successSummary"),
    pagesScanned: el("pagesScanned"), totalPages: el("totalPages"),
    linksFound: el("linksFound"), qrFound: el("qrFound"), processingTime: el("processingTime"), scanSpeed: el("scanSpeed"),
    linksList: el("linksList"), qrGrid: el("qrGrid"), linksEmpty: el("linksEmpty"), qrEmpty: el("qrEmpty"),
    urlPanel: el("urlPanel"), qrPanel: el("qrPanel"), urlPanelCount: el("urlPanelCount"), qrPanelCount: el("qrPanelCount"),
    progressUI: el("progressUI"), progressText: el("progressText"), progressPages: el("progressPages"),
    progressPct: el("progressPct"), progressFill: el("progressFill"), progressRing: document.querySelector(".scan-ring"),
    scanState: el("scanState"), fileMeta: el("fileMeta"), fileId: el("currentFileId"),
    liveUrls: el("liveUrls"), liveQr: el("liveQr"),
    error: el("errorMessage"), toast: el("toast"), search: el("resultSearch"), filter: el("resultFilter"),
    pageFilter: el("pageFilter"), sort: el("resultSort"), themeToggle: el("themeToggle"),
    dropZone: el("dropZone"), fileInput: el("pdfFile"), chooseBtn: el("chooseBtn")
};

let progressTimer = null;
let pollActive = false;
let isProcessing = false;
let activeRun = 0;
let activeController = null;
let lastProgressPct = 0;
let processStartedAt = 0;
let completionTimer = null;
let cachedLinks = [];
let cachedQrItems = [];

function clearResults() {
    ui.results.style.display = "none";
    ui.kpis.style.display = "none";
    ui.successBanner.style.display = "none";
    ui.pagesScanned.textContent = "0";
    ui.totalPages.textContent = "0";
    ui.linksFound.textContent = "0";
    ui.qrFound.textContent = "0";
    ui.processingTime.textContent = "0.0s";
    ui.scanSpeed.textContent = "0.0 pages/s";
    ui.liveUrls.textContent = "0";
    ui.liveQr.textContent = "0";
    ui.linksList.innerHTML = "";
    ui.qrGrid.innerHTML = "";
    ui.linksEmpty.style.display = "block";
    ui.qrEmpty.style.display = "block";
    ui.urlPanelCount.textContent = "0";
    ui.qrPanelCount.textContent = "0";
    ui.urlPanel.style.display = "block";
    ui.qrPanel.style.display = "block";
    ui.search.value = "";
    ui.filter.value = "all";
    ui.pageFilter.value = "all";
    ui.pageFilter.style.display = "none";
    ui.sort.value = "page";
    cachedLinks = [];
    cachedQrItems = [];
}

function toClickable(token) {
    if (!token) return "#";
    return token; // Preserve the extracted value exactly; never add a scheme.
}

function makeActionButton(label, icon, className) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `action-btn ${className || ""}`.trim();
    button.innerHTML = `<span aria-hidden="true">${icon}</span> ${label}`;
    return button;
}

function getDomain(token) {
    try { return new URL(toClickable(token)).hostname.replace(/^www\./, ""); }
    catch (_) { return ""; }
}

function getQrType(payload) {
    const value = String(payload || "").trim();
    if (/^WIFI:/i.test(value)) return {icon: "📶", label: "Wi-Fi", className: "type-wifi"};
    if (/^BEGIN:VCARD/i.test(value)) return {icon: "👤", label: "vCard", className: "type-vcard"};
    if (/^(upi:|upi:\/\/|pa=)/i.test(value)) return {icon: "💳", label: "UPI", className: "type-upi"};
    if (/^(https?:\/\/|www\.)/i.test(value)) return {icon: "🌐", label: "URL", className: "type-url"};
    return {icon: "📄", label: "Text", className: "type-text"};
}

const SOURCE_LABELS = {
    TEXT: "Native Text",
    PDF_LINK: "PDF Hyperlink",
    PDF_FORM: "PDF Form Field",
    QR: "QR",
    OCR: "OCR"
};

function formatSources(item) {
    const sources = typeof item === "object" ? (item.sources || [item.source || "TEXT"]) : ["TEXT"];
    return sources.map((source) => SOURCE_LABELS[source] || source).join(" • ");
}

// Enhanced: URL rows are responsive cards with page badges and delegated Copy/Open actions.
function renderLinks(linkHits) {
    ui.linksList.innerHTML = "";
    ui.urlPanelCount.textContent = String(linkHits.length);
    ui.linksEmpty.style.display = linkHits.length ? "none" : "block";

    for (const h of linkHits) {
        const token = typeof h === "string" ? h : h.token;
        const card = document.createElement("article");
        card.className = "result-card url-card";

        const content = document.createElement("div");
        content.className = "result-card-content";
        const top = document.createElement("div");
        top.className = "result-card-top";
        const badge = document.createElement("span");
        badge.className = "page-badge";
        // CHANGE: Keep only the PDF page number in the URL card badge.
        badge.textContent = typeof h === "object" && h.page ? `Page ${h.page}` : "URL";
        top.appendChild(badge);
        const value = document.createElement("p");
        value.className = "result-value";
        value.textContent = token;
        const domain = document.createElement("strong");
        domain.className = "domain-name";
        const domainName = getDomain(token);
        domain.textContent = domainName ? `🌐 ${domainName}` : "";
        content.appendChild(top);
        if (domainName) content.appendChild(domain);
        content.appendChild(value);

        const actions = document.createElement("div");
        actions.className = "result-actions";
        const copy = makeActionButton("Copy", "⧉");
        copy.dataset.copy = token;
        const open = document.createElement("a");
        open.className = "action-btn primary-action";
        open.href = toClickable(token);
        open.target = "_blank";
        open.rel = "noopener noreferrer";
        open.innerHTML = '<span aria-hidden="true">↗</span> Open';
        actions.append(copy, open);
        card.append(content, actions);
        ui.linksList.appendChild(card);
    }
}

// Enhanced: QR results use preview cards with payload, URL and consistent actions.
function renderQrItems(items) {
    ui.qrGrid.innerHTML = "";
    ui.qrPanelCount.textContent = String(items.length);
    ui.qrEmpty.style.display = items.length ? "none" : "block";

    for (const it of items) {
        const card = document.createElement("article");
        card.className = "result-card qr-result-card";
        if (it.preview) {
            const img = document.createElement("img");
            img.className = "qr-img";
            img.src = it.preview;
            img.alt = `QR code preview from page ${it.page}`;
            card.appendChild(img);
        }

        const content = document.createElement("div");
        content.className = "result-card-content";
        const top = document.createElement("div");
        top.className = "result-card-top";
        const badge = document.createElement("span");
        badge.className = "page-badge green-badge";
        const qrPages = Array.isArray(it.pages) && it.pages.length ? it.pages : [it.page];
        badge.textContent = `Page${qrPages.length > 1 ? "s" : ""} ${qrPages.join(", ")}`;
        const type = getQrType(it.payload);
        const typeBadge = document.createElement("span");
        typeBadge.className = `payload-type ${type.className}`;
        typeBadge.textContent = `${type.icon} ${type.label}`;
        top.append(badge, typeBadge);
        const payload = document.createElement("p");
        payload.className = "result-value payload";
        payload.textContent = it.payload || "";
        content.append(top, payload);

        if (it.url) {
            const url = document.createElement("a");
            url.className = "detected-url";
            url.href = toClickable(it.url);
            url.target = "_blank";
            url.rel = "noopener noreferrer";
            url.textContent = it.url;
            content.appendChild(url);
        }

        const actions = document.createElement("div");
        actions.className = "result-actions";
        const copy = makeActionButton("Copy payload", "⧉");
        copy.dataset.copy = it.payload || "";
        actions.appendChild(copy);
        if (it.preview) {
            const download = document.createElement("a");
            download.className = "action-btn";
            download.href = it.preview;
            download.download = `qr-page-${it.page}.png`;
            download.innerHTML = '<span aria-hidden="true">↓</span> Download QR';
            actions.appendChild(download);
        }
        if (it.url) {
            const open = document.createElement("a");
            open.className = "action-btn primary-action";
            open.href = toClickable(it.url);
            open.target = "_blank";
            open.rel = "noopener noreferrer";
            open.innerHTML = '<span aria-hidden="true">↗</span> Open';
            actions.appendChild(open);
        }
        content.appendChild(actions);
        card.appendChild(content);
        ui.qrGrid.appendChild(card);
    }
}

function applyFilters() {
    const query = ui.search.value.trim().toLowerCase();
    const mode = ui.filter.value;
    const selectedPage = ui.pageFilter.value;
    let links = cachedLinks.filter((h) => {
        const token = typeof h === "string" ? h : h.token;
        const matchesPage = mode !== "page" || selectedPage === "all" || String(h.page) === selectedPage;
        return String(token || "").toLowerCase().includes(query) && matchesPage;
    });
    let qrItems = cachedQrItems.filter((it) => {
        const matchesPage = mode !== "page" || selectedPage === "all" || String(it.page) === selectedPage;
        return `${it.payload || ""} ${it.url || ""}`.toLowerCase().includes(query) && matchesPage;
    });
    const sortMode = ui.sort.value;
    const pageOf = (item) => Number(typeof item === "string" ? 0 : item.page) || 0;
    const textOfLink = (item) => String(typeof item === "string" ? item : item.token || "").toLowerCase();
    const textOfQr = (item) => String(item.payload || "").toLowerCase();
    if (sortMode === "alpha") {
        links.sort((a, b) => textOfLink(a).localeCompare(textOfLink(b)));
        qrItems.sort((a, b) => textOfQr(a).localeCompare(textOfQr(b)));
    } else {
        links.sort((a, b) => pageOf(a) - pageOf(b));
        qrItems.sort((a, b) => pageOf(a) - pageOf(b));
    }
    ui.urlPanel.style.display = mode === "qr" ? "none" : "block";
    ui.qrPanel.style.display = mode === "urls" ? "none" : "block";
    renderLinks(links);
    renderQrItems(qrItems);
}

function populatePageFilter(totalPages) {
    ui.pageFilter.innerHTML = '<option value="all">All pages</option>';
    for (let page = 1; page <= totalPages; page++) {
        const option = document.createElement("option");
        option.value = String(page);
        option.textContent = `Page ${page}`;
        ui.pageFilter.appendChild(option);
    }
}

function stopProgressPoller() {
    pollActive = false;
    if (progressTimer !== null) clearTimeout(progressTimer);
    progressTimer = null;
}

function setRingProgress(percent) {
    ui.progressFill.style.strokeDasharray = String(RING_CIRCUMFERENCE);
    ui.progressFill.style.strokeDashoffset = String(RING_CIRCUMFERENCE * (1 - percent / 100));
    ui.progressRing.setAttribute("aria-valuenow", String(percent));
}

function resetProgressUI() {
    if (completionTimer !== null) window.clearTimeout(completionTimer);
    completionTimer = null;
    lastProgressPct = 0;
    ui.progressUI.classList.remove("done");
    ui.progressUI.style.display = "none";
    ui.progressText.textContent = "Processing...";
    ui.progressPages.textContent = "0/0";
    ui.progressPct.textContent = "0%";
    ui.scanState.className = "status-badge scanning";
    ui.scanState.innerHTML = "<i></i> Scanning";
    setRingProgress(0);
}

function updateElapsedTime(finalValue) {
    const elapsed = processStartedAt ? (Date.now() - processStartedAt) / 1000 : 0;
    ui.processingTime.textContent = `${elapsed.toFixed(finalValue ? 1 : 0)}s`;
}

function setProgressUI(currentPage, totalPages, linksCount, qrCount) {
    const total = Math.max(0, Number(totalPages) || 0);
    const current = Math.min(total, Math.max(0, Number(currentPage) || 0));
    const calculated = total > 0 ? Math.round((current / total) * 100) : 0;
    lastProgressPct = Math.min(100, Math.max(lastProgressPct, calculated));
    ui.progressUI.style.display = "flex";
    ui.kpis.style.display = "grid";
    ui.totalPages.textContent = String(total);
    ui.liveUrls.textContent = String(Math.max(0, Number(linksCount) || 0));
    ui.liveQr.textContent = String(Math.max(0, Number(qrCount) || 0));
    ui.linksFound.textContent = ui.liveUrls.textContent;
    ui.qrFound.textContent = ui.liveQr.textContent;
    ui.progressPages.textContent = `Page ${current}/${total}`;
    ui.progressText.textContent = total > 0 ? `Scanning page ${Math.min(current + 1, total)} of ${total}` : "Preparing document...";
    ui.progressPct.textContent = `${lastProgressPct}%`;
    setRingProgress(lastProgressPct);
    updateElapsedTime(false);
}

function setCompletedUI(totalPages) {
    if (completionTimer !== null) window.clearTimeout(completionTimer);
    lastProgressPct = 100;
    ui.progressUI.style.display = "flex";
    ui.progressUI.classList.add("done");
    ui.progressText.textContent = "Scan completed";
    ui.progressPages.textContent = `Page ${totalPages}/${totalPages}`;
    ui.progressPct.textContent = "100%";
    ui.scanState.className = "status-badge completed";
    ui.scanState.innerHTML = "<i></i> Completed";
    setRingProgress(100);
    updateElapsedTime(true);
    // Simplified: briefly show success, then leave only actions, KPIs, and results.
    completionTimer = window.setTimeout(() => {
        ui.progressUI.style.display = "none";
        completionTimer = null;
    }, 650);
}

function showError(message) {
    ui.error.textContent = message;
    ui.error.style.display = "block";
}

function clearError() {
    ui.error.textContent = "";
    ui.error.style.display = "none";
}

function setUploadDisabled(disabled) {
    ui.chooseBtn.disabled = disabled;
    ui.fileInput.disabled = disabled;
    ui.dropZone.setAttribute("aria-disabled", String(disabled));
}

function showToast(message) {
    ui.toast.textContent = message;
    ui.toast.classList.add("show");
    window.setTimeout(() => ui.toast.classList.remove("show"), 1800);
}

function startProgressPoller(fileId, runId, signal) {
    stopProgressPoller();
    pollActive = true;
    const poll = async () => {
        if (!pollActive || runId !== activeRun) return;
        try {
            const response = await fetch(`/api/progress/${fileId}`, {cache: "no-store", signal});
            if (!response.ok) throw new Error("Progress request failed");
            const state = await response.json();
            if (!pollActive || runId !== activeRun) return;
            setProgressUI(state.current_page, state.total_pages, state.links_count, state.qr_count);
            if (state.done === true) {
                stopProgressPoller();
                setCompletedUI(Number(state.total_pages) || 0);
                return;
            }
            progressTimer = setTimeout(poll, 700);
        } catch (error) {
            stopProgressPoller();
            if (error.name !== "AbortError" && runId === activeRun) showError("Unable to update progress. Please check your connection.");
        }
    };
    poll();
}

async function readJson(response) {
    try { return await response.json(); } catch (_) { return {}; }
}

async function uploadAndProcess(file) {
    if (!file || isProcessing) return;
    if (!file.name.toLowerCase().endsWith(".pdf")) {
        showError("Please choose a PDF file.");
        return;
    }
    if (file.size > 30 * 1024 * 1024) {
        showError("The PDF must be 30 MB or smaller.");
        return;
    }

    const runId = ++activeRun;
    isProcessing = true;
    processStartedAt = Date.now();
    activeController = new AbortController();
    setUploadDisabled(true);
    stopProgressPoller();
    resetProgressUI();
    clearResults();
    clearError();
    ui.uploadSection.classList.add("collapsed");
    ui.fileMeta.style.display = "flex";
    ui.fileMeta.textContent = `📄 ${file.name}`;

    try {
        const form = new FormData();
        form.append("file", file);
        const uploadResponse = await fetch("/api/upload", {method: "POST", body: form, signal: activeController.signal});
        const uploadJson = await readJson(uploadResponse);
        if (!uploadResponse.ok || !uploadJson.file_id) throw new Error(uploadJson.error || "Upload failed. Please try again.");
        if (runId !== activeRun) return;
        ui.fileId.value = uploadJson.file_id;
        setProgressUI(0, 0, 0, 0);
        startProgressPoller(uploadJson.file_id, runId, activeController.signal);

        const processResponse = await fetch(`/api/process/${uploadJson.file_id}`, {cache: "no-store", signal: activeController.signal});
        const processJson = await readJson(processResponse);
        stopProgressPoller();
        if (!processResponse.ok) throw new Error(processJson.error || "Processing failed. Please try again.");
        if (runId !== activeRun) return;

        // CHANGE: Preserve backend physical occurrences without UI numbering.
        cachedLinks = processJson.links || [];
        cachedQrItems = processJson.qr_items || [];
        ui.pagesScanned.textContent = processJson.pages_scanned ?? 0;
        ui.totalPages.textContent = processJson.total_pages ?? 0;
        ui.linksFound.textContent = processJson.links_count ?? 0;
        ui.qrFound.textContent = processJson.qr_count ?? 0;
        populatePageFilter(Number(processJson.total_pages) || 0);
        applyFilters();
        ui.kpis.style.display = "grid";
        ui.results.style.display = "block";
        setCompletedUI(Number(processJson.total_pages) || 0);
        const elapsedSeconds = Math.max(.1, (Date.now() - processStartedAt) / 1000);
        ui.scanSpeed.textContent = `${((Number(processJson.total_pages) || 0) / elapsedSeconds).toFixed(1)} pages/s`;
        ui.successSummary.textContent = `Processed ${processJson.total_pages || 0} pages in ${elapsedSeconds.toFixed(1)} seconds.`;
        ui.successBanner.style.display = "flex";
        ui.results.scrollIntoView({behavior: "smooth", block: "start"});
    } catch (error) {
        stopProgressPoller();
        if (error.name !== "AbortError" && runId === activeRun) showError(error.message || "A network error occurred. Please try again.");
    } finally {
        if (runId === activeRun) {
            isProcessing = false;
            activeController = null;
            setUploadDisabled(false);
        }
    }
}

ui.chooseBtn.addEventListener("click", (event) => { event.stopPropagation(); if (!isProcessing) ui.fileInput.click(); });
ui.dropZone.addEventListener("click", () => { if (!isProcessing) ui.fileInput.click(); });
ui.dropZone.addEventListener("keydown", (event) => {
    if (!isProcessing && (event.key === "Enter" || event.key === " ")) { event.preventDefault(); ui.fileInput.click(); }
});
ui.fileInput.addEventListener("change", () => uploadAndProcess(ui.fileInput.files && ui.fileInput.files[0]));
ui.dropZone.addEventListener("dragover", (event) => { event.preventDefault(); if (!isProcessing) ui.dropZone.classList.add("dragover"); });
ui.dropZone.addEventListener("dragleave", () => ui.dropZone.classList.remove("dragover"));
ui.dropZone.addEventListener("drop", (event) => {
    event.preventDefault();
    ui.dropZone.classList.remove("dragover");
    if (!isProcessing) uploadAndProcess(event.dataTransfer.files && event.dataTransfer.files[0]);
});

ui.search.addEventListener("input", applyFilters);
ui.filter.addEventListener("change", () => { ui.pageFilter.style.display = ui.filter.value === "page" ? "block" : "none"; applyFilters(); });
ui.pageFilter.addEventListener("change", applyFilters);
ui.sort.addEventListener("change", applyFilters);
ui.results.addEventListener("click", async (event) => {
    const button = event.target.closest("[data-copy]");
    if (!button) return;
    try {
        await navigator.clipboard.writeText(button.dataset.copy);
        showToast("Copied!");
    } catch (_) {
        showToast("Unable to copy");
    }
});

// Enhanced: create a formatted Excel 2003 XML workbook with URL and QR worksheets.
el("exportBtn").addEventListener("click", () => {
    const xmlEscape = (value) => String(value ?? "")
        .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    const cell = (value, style) => `<Cell${style ? ` ss:StyleID="${style}"` : ""}><Data ss:Type="String">${xmlEscape(value)}</Data></Cell>`;
    const row = (values, style) => `<Row>${values.map((value) => cell(value, style)).join("")}</Row>`;
    const urlRows = cachedLinks.map((item) => row([
        typeof item === "string" ? "" : (item.pages || [item.page]).join(", "),
        typeof item === "string" ? item : item.token,
        formatSources(item)
    ]));
    const qrRows = cachedQrItems.map((item) => row([
        item.payload, item.url || "", (item.pages || [item.page]).join(", "), getQrType(item.payload).label
    ]));
    const workbook = `<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
<Styles><Style ss:ID="Default"><Alignment ss:Vertical="Top"/><Font ss:FontName="Calibri" ss:Size="11"/></Style><Style ss:ID="Header"><Font ss:Bold="1" ss:Color="#292929"/><Interior ss:Color="#F5C400" ss:Pattern="Solid"/></Style></Styles>
<Worksheet ss:Name="URL Results"><Table><Column ss:Width="100"/><Column ss:Width="360"/><Column ss:Width="160"/>${row(["Page Number", "URL", "Extraction Source(s)"], "Header")}${urlRows.join("")}</Table><WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><AutoFilter x:Range="R1C1:R${urlRows.length + 1}C3" xmlns:x="urn:schemas-microsoft-com:office:excel"/></WorksheetOptions></Worksheet>
<Worksheet ss:Name="QR Code Results"><Table><Column ss:Width="360"/><Column ss:Width="300"/><Column ss:Width="100"/><Column ss:Width="90"/>${row(["Payload", "Extracted URL", "Pages", "Type"], "Header")}${qrRows.join("")}</Table><WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><AutoFilter x:Range="R1C1:R${qrRows.length + 1}C4" xmlns:x="urn:schemas-microsoft-com:office:excel"/></WorksheetOptions></Worksheet>
</Workbook>`;
    const url = URL.createObjectURL(new Blob([workbook], {type: "application/vnd.ms-excel"}));
    const anchor = document.createElement("a");
    anchor.href = url; anchor.download = "pdf-scan-results.xls"; anchor.click();
    URL.revokeObjectURL(url);
});

el("scanAnotherBtn").addEventListener("click", () => el("clearBtn").click());
ui.themeToggle.addEventListener("click", () => {
    const dark = document.body.classList.toggle("dark-mode");
    ui.themeToggle.textContent = dark ? "☀" : "☾";
    ui.themeToggle.setAttribute("aria-label", dark ? "Use light mode" : "Use dark mode");
});

el("clearBtn").addEventListener("click", () => {
    activeRun++;
    stopProgressPoller();
    if (activeController) activeController.abort();
    activeController = null;
    isProcessing = false;
    processStartedAt = 0;
    setUploadDisabled(false);
    resetProgressUI();
    clearResults();
    clearError();
    ui.uploadSection.classList.remove("collapsed");
    ui.fileInput.value = "";
    ui.fileId.value = "";
    ui.fileMeta.textContent = "";
    ui.fileMeta.style.display = "none";
    ui.dropZone.classList.remove("dragover");
});
