const DEFAULT_CATEGORIES = {
    libur_nasional: { label: "Libur Nasional / Umum", color: "#ef4444", textColor: "white", defaultNonEffective: true },
    libur_sekolah: { label: "Libur Jeda Semester / Akhir Tahun", color: "#f97316", textColor: "white", defaultNonEffective: true },
    libur_khusus: { label: "Libur Awal Puasa / Hari Raya", color: "#fb923c", textColor: "black", defaultNonEffective: true },
    ppdb: { label: "PPDB / MPLS", color: "#8b5cf6", textColor: "white", defaultNonEffective: true },
    p5: { label: "Projek Penguatan Profil Pelajar Pancasila", color: "#14b8a6", textColor: "white", defaultNonEffective: true },
    anbk: { label: "ANBK / Asesmen Nasional", color: "#06b6d4", textColor: "white", defaultNonEffective: true },
    ujian_tengah: { label: "PTS / STS Tengah Semester", color: "#eab308", textColor: "black", defaultNonEffective: true },
    ujian_akhir: { label: "PAS / PAT / SAS / SAT Akhir Semester", color: "#ca8a04", textColor: "white", defaultNonEffective: true },
    ujian_sekolah: { label: "Ujian Sekolah / Asesmen Akhir Jenjang", color: "#854d0e", textColor: "white", defaultNonEffective: true },
    rapor: { label: "Pembagian Buku Laporan Pendidikan", color: "#ec4899", textColor: "white", defaultNonEffective: false },
    kegiatan_siswa: { label: "Kegiatan Kesiswaan", color: "#22c55e", textColor: "white", defaultNonEffective: false },
    rapat_guru: { label: "Rapat Dinas / IHT / Workshop Guru", color: "#64748b", textColor: "white", defaultNonEffective: false },
    lainnya: { label: "Kegiatan Sekolah Lainnya", color: "#3b82f6", textColor: "white", defaultNonEffective: false }
};

const INITIAL_START_YEAR = getCurrentAcademicStartYear();

const DEFAULT_SETTINGS = {
    schoolName: "SMP NEGERI 2 SECANG",
    startYear: INITIAL_START_YEAR,
    lastActiveStartYear: INITIAL_START_YEAR,
    institutionType: "dinas",
    workDays: 5,
    className: "Kelas 8",
    subjectName: "Ilmu Pengetahuan Alam",
    jamPerMinggu: 24,
    headmasterName: "SUHARDI, S.Pd., M.M.Pd.",
    headmasterNip: "19720624 199903 1 002",
    teacherName: "YUSUF ARIS SYAMSUDDIN",
    teacherNip: "19950430 202321 1 004",
    ttdPlace: "Secang",
    ttdDate: "2026-07-13"
};

const monthNames = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
];
const dayNames = ["Mgg", "Sen", "Sel", "Rab", "Kam", "Jum", "Sab"];
const dayNamesFull = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
const UI_STATE_KEY = "siwaka.uiState.v4";
const LEGACY_UI_STATE_KEYS = ["siwaka.uiState.v3", "siwaka.uiState.v2", "siwaka.uiState.v1"];
const LOCAL_DATA_KEY = "siwaka.localData.v2";
const LEGACY_LOCAL_DATA_KEYS = ["siwaka.localData.v1"];
const EXPORT_OPTIONS_KEY = "siwaka.exportOptions.v1";
const ACTIVE_YEAR_KEY = "siwaka.activeStartYear.v2";
const WINDOW_STATE_KEY = "siwaka.windowState.v1";
const SIDEBAR_STATE_KEY = "siwaka.sidebarCollapsed.v1";
const HANDLE_DB_NAME = "siwaka-file-handles";
const HANDLE_STORE_NAME = "handles";
const DATABASE_HANDLE_KEY = "kalender-database";
const DEFAULT_PDF_OPTIONS = {
    paper: "a3",
    orientation: "landscape",
    margin: 6,
    captureScale: 2,
    fitPercent: 100
};
const DEFAULT_WORD_OPTIONS = {
    paper: "a4",
    orientation: "landscape",
    margin: 8,
    fitMode: "compact"
};

let categories = cloneCategories(DEFAULT_CATEGORIES);
let database = {
    settings: { ...DEFAULT_SETTINGS },
    events: [],
    periods: [],
    categories
};
let settings = normalizeSettings(database.settings);
let periods = [];
let activeStartYear = settings.startYear;
let currentDate = new Date(activeStartYear, 6, 1);
let activeView = "dashboard";
let dashboardMode = "month";
let academicFormat = "model-b";
let activeSettingsTab = "database";
let sidebarCollapsed = false;
let activeFilters = new Set(Object.keys(categories));
let dbFileHandle = null;
let dbFileName = "kalender_database.xlsx";
let dbDirty = false;
let databaseMissingOnBoot = false;
let restoredFromLocalSnapshot = false;
let booting = true;
let autoSaveTimer = null;
let pdfPreviewRequestId = 0;

const $ = (selector, root = document) => root.querySelector(selector);
const $$ = (selector, root = document) => Array.from(root.querySelectorAll(selector));

document.addEventListener("DOMContentLoaded", init);

async function init() {
    bindStaticEvents();
    const savedState = loadUiState();
    await recoverDatabaseHandle();
    let loaded = false;
    let loadedSource = "";
    if (dbFileHandle) {
        loaded = await loadDatabaseFromHandle(false);
        if (loaded) loadedSource = "handle";
    }
    if (!loaded) {
        loaded = await loadBundledDatabase();
        if (loaded) loadedSource = "bundled";
    }
    if (loadedSource === "bundled") {
        // Only hydrate from local if it's actually newer than what we have in bundled
        hydrateFromLocalSnapshotIfNewer();
    }
    if (!loaded) {
        loaded = loadLocalDataSnapshot();
        if (loaded) loadedSource = "local";
    }
    if (!loaded) {
        database.settings = normalizeSettings(DEFAULT_SETTINGS);
        database.categories = cloneCategories(DEFAULT_CATEGORIES);
        database.events = [];
        database.periods = [];
    }

    settings = normalizeSettings(database.settings);
    categories = normalizeCategories(database.categories);
    database.categories = categories;
    database.events = normalizeEvents(database.events);
    database.periods = normalizePeriods(database.periods, database.events, settings.startYear);
    periods = database.periods;
    settings.startYear = getDefaultPeriodYear();
    database.settings = { ...settings };
    activeFilters = new Set(Object.keys(categories));
    restoreUiState(savedState);

    setupYearSelectors();
    setupFilters();
    setupCategoryOptions();
    populateRekapMonthFilter();
    restoreUiControls(savedState);
    populateSettingsForm();
    applySettingsToDOM();
    resetCategoryForm();
    renderCategoryTable();
    renderPeriodTable();
    switchSettingsTab(activeSettingsTab);
    switchView(activeView || "dashboard");
    booting = false;
    setDatabaseDirty(false);
    saveUiState();
    await attemptDatabaseRegeneration();
}

function loadUiState() {
    const states = [UI_STATE_KEY, ...LEGACY_UI_STATE_KEYS]
        .map((key) => {
            try {
                const value = JSON.parse(localStorage.getItem(key) || "null");
                return value && typeof value === "object" ? { ...value, __key: key } : null;
            } catch (error) {
                return null;
            }
        })
        .filter(Boolean);
    const newest = states.sort((a, b) => getStateTimestamp(b) - getStateTimestamp(a))[0] || {};
    const bootState = window.__SIWAKA_BOOT_STATE__ && typeof window.__SIWAKA_BOOT_STATE__ === "object"
        ? window.__SIWAKA_BOOT_STATE__
        : {};
    const storedActiveYear = validStartYear(safeStorageGet(localStorage, ACTIVE_YEAR_KEY))
        || validStartYear(safeStorageGet(sessionStorage, ACTIVE_YEAR_KEY))
        || validStartYear(bootState.activeStartYear);
    return {
        ...newest,
        ...bootState,
        activeStartYear: storedActiveYear || newest.activeStartYear || bootState.activeStartYear
    };
}

function getStateTimestamp(state = {}) {
    const timestamp = Date.parse(state.savedAt || "");
    return Number.isFinite(timestamp) ? timestamp : 0;
}

function loadLocalDataSnapshot() {
    const snapshot = readLocalDataSnapshot();
    if (!snapshot) return false;
    applyDataSnapshot(snapshot);
    restoredFromLocalSnapshot = true;
    if (snapshot.__key !== LOCAL_DATA_KEY) saveLocalDataSnapshot();
    return true;
}

function readLocalDataSnapshot() {
    for (const key of [LOCAL_DATA_KEY, ...LEGACY_LOCAL_DATA_KEYS]) {
        const snapshot = readLocalDataSnapshotByKey(key);
        if (snapshot) return { ...snapshot, __key: key };
    }
    return null;
}

function readLocalDataSnapshotByKey(key) {
    try {
        const snapshot = JSON.parse(localStorage.getItem(key) || "null");
        if (!snapshot || !snapshot.settings || !Array.isArray(snapshot.events)) return null;
        if (validStartYear(snapshot.activeStartYear)) snapshot.settings.lastActiveStartYear = snapshot.activeStartYear;
        return snapshot;
    } catch (error) {
        console.warn("Snapshot lokal tidak dapat dimuat.", error);
        return null;
    }
}

function applyDataSnapshot(snapshot) {
    database.settings = normalizeSettings(snapshot.settings);
    settings = normalizeSettings(database.settings);
    const snapshotCategories = normalizeCategories(snapshot.categories || DEFAULT_CATEGORIES);
    categories = snapshotCategories;
    database.categories = snapshotCategories;
    database.events = normalizeEvents(snapshot.events);
    database.periods = normalizePeriods(snapshot.periods || [], database.events, settings.startYear);
    periods = database.periods;
    settings.startYear = getDefaultPeriodYear();
    database.settings = { ...settings };
}

function hydrateFromLocalSnapshotIfNewer() {
    const snapshot = readLocalDataSnapshotByKey(LOCAL_DATA_KEY);
    if (!snapshot) return false;
    const localTime = Date.parse(snapshot.savedAt || snapshot.settings?.savedAt || "");
    const databaseTime = Date.parse(database.settings?.savedAt || "");
    if (Number.isFinite(localTime) && (!Number.isFinite(databaseTime) || localTime >= databaseTime)) {
        applyDataSnapshot(snapshot);
        restoredFromLocalSnapshot = true;
        return true;
    }
    return false;
}

function restoreUiState(state = {}) {
    activeStartYear = resolveActiveStartYear(state);
    currentDate = normalizeSavedDate(state.currentDate);
    if (!currentDate || !dateInAcademicYear(currentDate, activeStartYear)) currentDate = new Date(activeStartYear, 6, 1);
    activeView = ["dashboard", "akademik", "efektif", "rekap", "settings"].includes(state.activeView) ? state.activeView : "dashboard";
    dashboardMode = ["month", "semester1", "semester2", "year"].includes(state.dashboardMode) ? state.dashboardMode : "month";
    academicFormat = ["model-b", "model-c", "all"].includes(state.academicFormat) ? state.academicFormat : "model-b";
    activeSettingsTab = ["database", "identity", "periods", "categories"].includes(state.activeSettingsTab) ? state.activeSettingsTab : "database";
    sidebarCollapsed = Boolean(state.sidebarCollapsed) || safeStorageGet(localStorage, SIDEBAR_STATE_KEY) === "true";
    if (Array.isArray(state.activeFilters)) {
        const valid = state.activeFilters.filter((key) => categories[key]);
        if (valid.length) activeFilters = new Set(valid);
    }
}

function restoreUiControls(state = {}) {
    applySidebarState();
    $$(".segment[data-dashboard-mode]").forEach((button) => button.classList.toggle("active", button.dataset.dashboardMode === dashboardMode));
    $$(".segment[data-academic-format]").forEach((button) => button.classList.toggle("active", button.dataset.academicFormat === academicFormat));
    if (state.academicScope && $("#academic-scope").querySelector(`option[value="${state.academicScope}"]`)) $("#academic-scope").value = state.academicScope;
    if (state.rekapSemester && $("#filter-rekap-semester").querySelector(`option[value="${state.rekapSemester}"]`)) $("#filter-rekap-semester").value = state.rekapSemester;
    if (state.rekapMonth && $("#filter-rekap-bulan").querySelector(`option[value="${state.rekapMonth}"]`)) $("#filter-rekap-bulan").value = state.rekapMonth;
}

function saveUiState() {
    if (booting) return;
    const state = {
        activeView,
        dashboardMode,
        academicFormat,
        activeSettingsTab,
        activeStartYear,
        currentDate: formatDate(currentDate),
        activeFilters: Array.from(activeFilters),
        academicScope: $("#academic-scope")?.value || "year",
        rekapSemester: $("#filter-rekap-semester")?.value || "all",
        rekapMonth: $("#filter-rekap-bulan")?.value || "all",
        sidebarCollapsed,
        savedAt: new Date().toISOString()
    };
    try {
        localStorage.setItem(UI_STATE_KEY, JSON.stringify(state));
        localStorage.setItem(SIDEBAR_STATE_KEY, sidebarCollapsed ? "true" : "false");
        persistActiveYear(activeStartYear);
        saveLocalDataSnapshot();
    } catch (error) {
        console.warn("Preferensi tampilan tidak dapat disimpan.", error);
    }
}

function resolveActiveStartYear(state = {}) {
    return validStartYear(state.activeStartYear)
        || validStartYear(settings.lastActiveStartYear)
        || getDefaultPeriodYear()
        || validStartYear(settings.startYear)
        || inferActiveStartYearFromEvents()
        || DEFAULT_SETTINGS.startYear;
}

function inferActiveStartYearFromEvents() {
    const years = (database.events || [])
        .flatMap((event) => [event.start, event.end])
        .map((value) => inferAcademicYear(value))
        .filter((year) => validStartYear(year));
    if (!years.length) return 0;
    return Math.max(...years);
}

function safeStorageGet(storage, key) {
    try {
        return storage.getItem(key);
    } catch (error) {
        return "";
    }
}

function validStartYear(value) {
    const year = Number(value);
    return Number.isInteger(year) && year >= 2000 && year <= 2100 ? year : 0;
}

function loadWindowUiState() {
    try {
        const parsed = JSON.parse(window.name || "{}");
        return parsed && parsed[WINDOW_STATE_KEY] ? parsed[WINDOW_STATE_KEY] : {};
    } catch (error) {
        return {};
    }
}

function saveWindowUiStatePatch(patch) {
    try {
        const parsed = JSON.parse(window.name || "{}");
        parsed[WINDOW_STATE_KEY] = { ...(parsed[WINDOW_STATE_KEY] || {}), ...patch };
        window.name = JSON.stringify(parsed);
    } catch (error) {
        window.name = JSON.stringify({ [WINDOW_STATE_KEY]: patch });
    }
}

function persistActiveYear(year) {
    const normalized = validStartYear(year) || DEFAULT_SETTINGS.startYear;
    activeStartYear = normalized;
    settings.lastActiveStartYear = normalized;
    database.settings = { ...database.settings, ...settings, lastActiveStartYear: normalized };
    try {
        localStorage.setItem(ACTIVE_YEAR_KEY, String(normalized));
        sessionStorage.setItem(ACTIVE_YEAR_KEY, String(normalized));
    } catch (error) {
        console.warn("Periode aktif tidak dapat disimpan di localStorage.", error);
    }
    saveWindowUiStatePatch({ activeStartYear: normalized, savedAt: new Date().toISOString() });
}

function saveLocalDataSnapshot() {
    try {
        const savedAt = new Date().toISOString();
        localStorage.setItem(LOCAL_DATA_KEY, JSON.stringify({
            savedAt,
            activeStartYear,
            settings: { ...settings, lastActiveStartYear: activeStartYear, savedAt },
            events: database.events,
            periods,
            categories
        }));
    } catch (error) {
        console.warn("Snapshot data lokal tidak dapat disimpan.", error);
    }
}

function normalizeSavedDate(value) {
    if (!value) return null;
    const date = parseDate(value);
    return Number.isNaN(date.getTime()) ? null : date;
}

function dateInAcademicYear(date, startYear) {
    const start = new Date(Number(startYear), 6, 1);
    const end = new Date(Number(startYear) + 1, 5, 30);
    return date >= start && date <= end;
}

function bindStaticEvents() {
    window.addEventListener("beforeunload", () => saveUiState());
    window.addEventListener("resize", applySidebarState);
    document.addEventListener("visibilitychange", () => {
        if (document.visibilityState === "hidden") saveUiState();
    });
    $("#btn-open-sidebar").addEventListener("click", toggleSidebar);
    $("#btn-close-sidebar").addEventListener("click", toggleSidebar);
    $("#mobile-overlay").addEventListener("click", closeSidebar);
    $$(".nav-item").forEach((button) => {
        button.addEventListener("click", () => switchView(button.dataset.view));
    });

    const handleYearSelectorChange = (event) => {
        activeStartYear = validStartYear(event.target.value) || activeStartYear || settings.startYear;
        persistActiveYear(activeStartYear);
        currentDate = new Date(activeStartYear, 6, 1);
        applySettingsToDOM();
        populateRekapMonthFilter();
        saveUiState();
        markDataChanged();
        renderCurrentView();
    };
    $("#header-year-selector").addEventListener("change", handleYearSelectorChange);
    $("#header-year-selector").addEventListener("input", handleYearSelectorChange);

    $("#header-actions").addEventListener("click", handleHeaderAction);
    $("#btn-prev-month").addEventListener("click", () => changeMonth(-1));
    $("#btn-next-month").addEventListener("click", () => changeMonth(1));
    $("#btn-today").addEventListener("click", goToday);
    $("#btn-select-all-filters").addEventListener("click", selectAllFilters);
    $("#academic-scope").addEventListener("change", () => {
        saveUiState();
        renderAkademikView();
    });
    $("#filter-rekap-semester").addEventListener("change", () => {
        saveUiState();
        renderRekapView();
    });
    $("#filter-rekap-bulan").addEventListener("change", () => {
        saveUiState();
        renderRekapView();
    });
    $("#export-dashboard").addEventListener("click", handleReportClick);
    $("#view-rekap").addEventListener("click", handleReportClick);

    $$(".segment[data-dashboard-mode]").forEach((button) => {
        button.addEventListener("click", () => {
            dashboardMode = button.dataset.dashboardMode;
            $$(".segment[data-dashboard-mode]").forEach((item) => item.classList.toggle("active", item === button));
            saveUiState();
            renderDashboardView();
        });
    });

    $$(".segment[data-academic-format]").forEach((button) => {
        button.addEventListener("click", () => {
            academicFormat = button.dataset.academicFormat;
            $$(".segment[data-academic-format]").forEach((item) => item.classList.toggle("active", item === button));
            saveUiState();
            renderAkademikView();
        });
    });

    $$(".segment[data-settings-tab]").forEach((button) => {
        button.addEventListener("click", () => switchSettingsTab(button.dataset.settingsTab));
    });

    $("#event-form").addEventListener("submit", handleEventSubmit);
    $("#btn-close-modal").addEventListener("click", closeEventModal);
    $("#btn-cancel-event").addEventListener("click", closeEventModal);
    $("#btn-delete-event").addEventListener("click", () => deleteEvent($("#event-id").value));
    $("#event-start").addEventListener("change", () => {
        const start = $("#event-start").value;
        $("#event-end").min = start;
        if ($("#event-end").value < start) $("#event-end").value = start;
    });
    $("#event-type").addEventListener("change", () => {
        const category = categories[$("#event-type").value];
        if (category) $("#event-isNonEffective").checked = Boolean(category.defaultNonEffective);
    });

    $("#settings-form").addEventListener("submit", saveSettings);
    $("#set-institutionType").addEventListener("change", handleInstitutionChange);
    $("#set-workDays").addEventListener("change", handleWorkDaysChange);
    $("#btn-open-database").addEventListener("click", openDatabase);
    $("#btn-save-database").addEventListener("click", () => saveDatabase({ manual: true }));
    $("#btn-download-database").addEventListener("click", downloadDatabase);
    $("#db-file-input").addEventListener("change", loadDatabaseFromInput);

    $("#category-form").addEventListener("submit", saveCategory);
    $("#btn-new-category").addEventListener("click", resetCategoryForm);
    $("#btn-reset-category").addEventListener("click", resetCategoryForm);
    $("#table-categories").addEventListener("click", handleCategoryTableClick);
    $("#btn-add-period").addEventListener("click", addLearningPeriod);
    $("#table-periods").addEventListener("click", handlePeriodTableClick);
}

function toggleSidebar() {
    if (window.innerWidth <= 820) {
        document.body.classList.toggle("sidebar-open");
    } else {
        sidebarCollapsed = !document.body.classList.contains("sidebar-collapsed");
        applySidebarState();
        saveUiState();
    }
}

function closeSidebar() {
    document.body.classList.remove("sidebar-open");
}

function applySidebarState() {
    document.body.classList.toggle("sidebar-collapsed", sidebarCollapsed && window.innerWidth > 820);
}

function switchSettingsTab(tab) {
    activeSettingsTab = tab;
    $$(".segment[data-settings-tab]").forEach((button) => button.classList.toggle("active", button.dataset.settingsTab === tab));
    $$(".settings-tab-panel").forEach((panel) => panel.classList.toggle("active", panel.id === `settings-tab-${tab}`));
    saveUiState();
}

function setupFilters() {
    const container = $("#filter-container");
    container.innerHTML = Object.entries(categories).map(([key, category]) => `
        <label class="filter-item" title="${escapeHtml(category.label)}">
            <input type="checkbox" data-filter="${key}" ${activeFilters.has(key) ? "checked" : ""}>
            <span class="swatch" style="background:${category.color}"></span>
            <span>${escapeHtml(category.label)}</span>
        </label>
    `).join("");

    container.onchange = (event) => {
        const key = event.target.dataset.filter;
        if (!key) return;
        if (event.target.checked) activeFilters.add(key);
        else activeFilters.delete(key);
        saveUiState();
        renderCurrentView();
    };
}

function setupCategoryOptions() {
    $("#event-type").innerHTML = Object.entries(categories).map(([key, category]) => (
        `<option value="${key}">${escapeHtml(category.label)}</option>`
    )).join("");
}

function setupYearSelectors() {
    activeStartYear = validStartYear(activeStartYear) || resolveActiveStartYear(loadUiState());
    const years = new Set([
        Number(settings.startYear),
        Number(activeStartYear),
        Number(settings.lastActiveStartYear),
        getCurrentAcademicStartYear()
    ]);
    periods.forEach((period) => years.add(Number(period.startYear)));
    database.events.forEach((event) => {
        const start = parseDate(event.start);
        const end = parseDate(event.end);
        if (!Number.isNaN(start.getTime())) {
            years.add(start.getMonth() >= 6 ? start.getFullYear() : start.getFullYear() - 1);
        }
        if (!Number.isNaN(end.getTime())) {
            years.add(end.getMonth() >= 6 ? end.getFullYear() : end.getFullYear() - 1);
        }
    });
    const options = Array.from(years).filter((year) => validStartYear(year)).sort((a, b) => a - b).map((year) => (
        `<option value="${year}">${year}/${year + 1}</option>`
    )).join("");
    $("#header-year-selector").innerHTML = options;
    $("#set-startYear").innerHTML = options;
    forceSelectValue($("#header-year-selector"), activeStartYear);
    forceSelectValue($("#set-startYear"), settings.startYear);
}

function forceSelectValue(select, year) {
    const normalized = validStartYear(year) || DEFAULT_SETTINGS.startYear;
    if (!select.querySelector(`option[value="${normalized}"]`)) {
        select.insertAdjacentHTML("beforeend", `<option value="${normalized}">${normalized}/${normalized + 1}</option>`);
    }
    select.value = String(normalized);
    if (select.value !== String(normalized)) {
        Array.from(select.options).forEach((option) => { option.selected = option.value === String(normalized); });
    }
}

function populateRekapMonthFilter() {
    const previous = $("#filter-rekap-bulan")?.value || "all";
    $("#filter-rekap-bulan").innerHTML = [
        `<option value="all">Semua Bulan</option>`,
        ...getAcademicMonths("year").map((item) => (
            `<option value="${item.y}-${String(item.m + 1).padStart(2, "0")}">${monthNames[item.m]} ${item.y}</option>`
        ))
    ].join("");
    if ($("#filter-rekap-bulan").querySelector(`option[value="${previous}"]`)) $("#filter-rekap-bulan").value = previous;
}

function selectAllFilters() {
    activeFilters = new Set(Object.keys(categories));
    $$("[data-filter]").forEach((input) => { input.checked = true; });
    saveUiState();
    renderCurrentView();
}

function switchView(viewId) {
    activeView = viewId;
    closeSidebar();
    $$(".view-section").forEach((section) => section.classList.toggle("active", section.id === `view-${viewId}`));
    $$(".nav-item").forEach((button) => button.classList.toggle("active", button.dataset.view === viewId));
    const titles = {
        dashboard: "Dashboard Kalender",
        akademik: "Format Akademik",
        efektif: "Waktu Efektif",
        rekap: "Rekapitulasi",
        settings: "Pengaturan Sistem"
    };
    $("#header-title").textContent = titles[viewId] || "Kalender Pendidikan";
    renderHeaderActions(viewId);
    saveUiState();
    renderCurrentView();
}

function renderHeaderActions(viewId) {
    const actionSets = {
        dashboard: `
            <button type="button" class="primary-button" data-action="new-event"><i class="fa-solid fa-plus"></i> Input</button>
            <button type="button" class="export-button" data-action="excel" data-target="export-dashboard" data-name="Dashboard_Kalender"><i class="fa-solid fa-file-excel"></i> Excel</button>
            <button type="button" class="export-button" data-action="pdf" data-target="export-dashboard" data-name="Dashboard_Kalender" data-orientation="landscape"><i class="fa-solid fa-file-pdf"></i> PDF</button>
            <button type="button" class="export-button" data-action="word" data-target="export-dashboard" data-name="Dashboard_Kalender"><i class="fa-solid fa-file-word"></i> Word A4</button>
        `,
        akademik: `
            <button type="button" class="export-button" data-action="excel" data-target="export-akademik" data-name="Format_Akademik"><i class="fa-solid fa-file-excel"></i> Excel</button>
            <button type="button" class="export-button" data-action="pdf" data-target="export-akademik" data-name="Format_Akademik" data-orientation="landscape"><i class="fa-solid fa-file-pdf"></i> PDF</button>
            <button type="button" class="export-button" data-action="word" data-target="export-akademik" data-name="Format_Akademik"><i class="fa-solid fa-file-word"></i> Word A4</button>
        `,
        efektif: `
            <button type="button" class="export-button" data-action="excel" data-target="export-efektif" data-name="Waktu_Efektif"><i class="fa-solid fa-file-excel"></i> Excel</button>
            <button type="button" class="export-button" data-action="pdf" data-target="export-efektif" data-name="Waktu_Efektif" data-orientation="portrait"><i class="fa-solid fa-file-pdf"></i> PDF</button>
            <button type="button" class="export-button" data-action="word" data-target="export-efektif" data-name="Waktu_Efektif"><i class="fa-solid fa-file-word"></i> Word A4</button>
        `,
        rekap: `
            <button type="button" class="primary-button" data-action="new-event"><i class="fa-solid fa-plus"></i> Input</button>
            <button type="button" class="secondary-button" data-action="copy-period"><i class="fa-solid fa-copy"></i> Salin Periode</button>
            <button type="button" class="export-button" data-action="excel" data-target="export-rekap" data-name="Rekapitulasi"><i class="fa-solid fa-file-excel"></i> Excel</button>
            <button type="button" class="export-button" data-action="pdf" data-target="export-rekap" data-name="Rekapitulasi" data-orientation="portrait"><i class="fa-solid fa-file-pdf"></i> PDF</button>
            <button type="button" class="export-button" data-action="word" data-target="export-rekap" data-name="Rekapitulasi"><i class="fa-solid fa-file-word"></i> Word A4</button>
        `,
        settings: ""
    };
    $("#header-actions").innerHTML = actionSets[viewId] || "";
}

function handleHeaderAction(event) {
    const button = event.target.closest("[data-action]");
    if (!button) return;
    const action = button.dataset.action;
    if (action === "new-event") openEventModal();
    if (action === "copy-period") openCopyPeriodDialog();
    if (action === "pdf") exportPDF(button.dataset.target, button.dataset.name, button.dataset.orientation || "landscape");
    if (action === "word") exportWord(button.dataset.target, button.dataset.name);
    if (action === "excel") exportExcel(button.dataset.target, button.dataset.name);
}

function renderCurrentView() {
    if (activeView === "dashboard") renderDashboardView();
    if (activeView === "akademik") renderAkademikView();
    if (activeView === "efektif") renderEfektifView();
    if (activeView === "rekap") renderRekapView();
    if (activeView === "settings") {
        renderCategoryTable();
        renderPeriodTable();
    }
}

function applySettingsToDOM() {
    const academicYear = `${activeStartYear}-${Number(activeStartYear) + 1}`;
    setText(".school-name-text", settings.schoolName);
    setText(".academic-year-text", academicYear);
    setText(".headmaster-name-text", settings.headmasterName);
    setText(".headmaster-nip-text", settings.headmasterNip);
    setText(".teacher-name-text", settings.teacherName);
    setText(".teacher-nip-text", settings.teacherNip);
    setText(".subject-name-text", settings.subjectName);
    setText(".class-name-text", settings.className);
    setText(".work-days-text", settings.workDays);
    setText(".jam-per-minggu-text", settings.jamPerMinggu);
    setText(".ttd-place-date", `${settings.ttdPlace}, ${formatDateIndo(settings.ttdDate)}`);
    forceSelectValue($("#header-year-selector"), activeStartYear);
}

function setText(selector, value) {
    $$(selector).forEach((node) => { node.textContent = value || ""; });
}

function populateSettingsForm() {
    setupYearSelectors();
    $("#set-schoolName").value = settings.schoolName || "";
    $("#set-startYear").value = String(settings.startYear);
    $("#set-institutionType").value = settings.institutionType || "custom";
    $("#set-workDays").value = String(settings.workDays || 5);
    $("#set-subjectName").value = settings.subjectName || "";
    $("#set-className").value = settings.className || "";
    $("#set-jamPerMinggu").value = settings.jamPerMinggu || 24;
    $("#set-headmasterName").value = settings.headmasterName || "";
    $("#set-headmasterNip").value = settings.headmasterNip || "";
    $("#set-teacherName").value = settings.teacherName || "";
    $("#set-teacherNip").value = settings.teacherNip || "";
    $("#set-ttdPlace").value = settings.ttdPlace || "";
    $("#set-ttdDate").value = settings.ttdDate || "";
}

function handleInstitutionChange() {
    const type = $("#set-institutionType").value;
    if (type === "dinas") $("#set-workDays").value = "5";
    if (type === "kemenag") $("#set-workDays").value = "6";
}

function handleWorkDaysChange() {
    const workDays = Number($("#set-workDays").value);
    const type = $("#set-institutionType").value;
    if ((type === "dinas" && workDays !== 5) || (type === "kemenag" && workDays !== 6)) {
        $("#set-institutionType").value = "custom";
    }
}

function saveSettings(event) {
    event.preventDefault();
    settings = normalizeSettings({
        schoolName: $("#set-schoolName").value.trim(),
        startYear: Number($("#set-startYear").value),
        institutionType: $("#set-institutionType").value,
        workDays: Number($("#set-workDays").value),
        className: $("#set-className").value.trim(),
        subjectName: $("#set-subjectName").value.trim(),
        jamPerMinggu: Number($("#set-jamPerMinggu").value),
        lastActiveStartYear: activeStartYear,
        headmasterName: $("#set-headmasterName").value.trim(),
        headmasterNip: $("#set-headmasterNip").value.trim(),
        teacherName: $("#set-teacherName").value.trim(),
        teacherNip: $("#set-teacherNip").value.trim(),
        ttdPlace: $("#set-ttdPlace").value.trim(),
        ttdDate: $("#set-ttdDate").value
    });
    ensurePeriod(settings.startYear, true);
    setDefaultPeriod(settings.startYear, false);
    activeStartYear = settings.startYear;
    database.settings = { ...settings };
    setupYearSelectors();
    applySettingsToDOM();
    populateRekapMonthFilter();
    renderPeriodTable();
    saveUiState();
    markDataChanged();
    renderCurrentView();
    toast("Pengaturan disimpan");
}

function normalizeSettings(value = {}) {
    const merged = { ...DEFAULT_SETTINGS, ...value };
    merged.startYear = Number(merged.startYear) || DEFAULT_SETTINGS.startYear;
    merged.lastActiveStartYear = validStartYear(merged.lastActiveStartYear) || "";
    merged.workDays = Number(merged.workDays) === 6 ? 6 : 5;
    merged.jamPerMinggu = Math.max(1, Number(merged.jamPerMinggu) || DEFAULT_SETTINGS.jamPerMinggu);
    if (!["dinas", "kemenag", "custom"].includes(merged.institutionType)) merged.institutionType = "custom";
    return merged;
}

function cloneCategories(source) {
    return JSON.parse(JSON.stringify(source));
}

function normalizeCategories(source = {}) {
    const result = {};
    Object.entries(source).forEach(([rawKey, raw]) => {
        const key = sanitizeCategoryKey(raw.key || rawKey);
        if (!key) return;
        const hasDefaultStatus = ["defaultNonEffective", "default_non_effective", "defaultTidakEfektif", "tidakEfektif", "isNonEffective"]
            .some((field) => Object.prototype.hasOwnProperty.call(raw, field));
        result[key] = {
            label: String(raw.label || raw.nama || raw.name || key).trim(),
            color: normalizeColor(raw.color || raw.warna || "#3b82f6"),
            textColor: String(raw.textColor || raw.text_color || raw.warnaTeks || "white").toLowerCase() === "black" ? "black" : "white",
            defaultNonEffective: hasDefaultStatus
                ? normalizeBoolean(raw.defaultNonEffective ?? raw.default_non_effective ?? raw.defaultTidakEfektif ?? raw.tidakEfektif ?? raw.isNonEffective)
                : Boolean((DEFAULT_CATEGORIES[key] || {}).defaultNonEffective)
        };
    });
    if (!Object.keys(result).length) return cloneCategories(DEFAULT_CATEGORIES);
    if (!result.lainnya) result.lainnya = { ...DEFAULT_CATEGORIES.lainnya };
    return result;
}

function normalizeEvents(events = []) {
    return events.map((event) => ({
        id: String(event.id || createId()),
        academicYear: event.academicYear ? Number(event.academicYear) : inferAcademicYear(event.start),
        title: String(event.title || "").trim(),
        start: normalizeDateCell(event.start),
        end: normalizeDateCell(event.end || event.start),
        type: categories[event.type] ? event.type : "lainnya",
        isNonEffective: Boolean(event.isNonEffective),
        notes: String(event.notes || "")
    })).filter((event) => event.title && event.start && event.end);
}

function normalizePeriods(source = [], events = database.events, defaultStartYear = settings.startYear) {
    const now = new Date().toISOString();
    const map = new Map();
    const add = (year, raw = {}) => {
        const startYear = validStartYear(year);
        if (!startYear) return;
        const existing = map.get(startYear) || {};
        map.set(startYear, {
            startYear,
            label: raw.label || existing.label || `${startYear}/${startYear + 1}`,
            isDefault: Boolean(raw.isDefault) || Boolean(existing.isDefault),
            createdAt: raw.createdAt || existing.createdAt || now,
            updatedAt: raw.updatedAt || existing.updatedAt || raw.createdAt || now
        });
    };
    (source || []).forEach((period) => add(period.startYear || period.tahunMulai || period.Tahun || period.year, {
        label: period.label || period.nama || period.Nama,
        isDefault: normalizeBoolean(period.isDefault ?? period.default ?? period.Default),
        createdAt: period.createdAt || period.dibuat,
        updatedAt: period.updatedAt || period.diubah
    }));
    add(defaultStartYear, { isDefault: true });
    (events || []).forEach((event) => {
        add(inferAcademicYear(event.start));
        add(inferAcademicYear(event.end));
    });
    const defaultYear = validStartYear(defaultStartYear) || Array.from(map.keys()).sort((a, b) => a - b)[0] || getCurrentAcademicStartYear();
    if (!map.has(defaultYear)) add(defaultYear, { isDefault: true });
    return Array.from(map.values())
        .map((period) => ({ ...period, isDefault: period.startYear === defaultYear }))
        .sort((a, b) => a.startYear - b.startYear);
}

function getDefaultPeriodYear() {
    const periodDefault = periods.find((period) => period.isDefault);
    return validStartYear(settings.startYear)
        || validStartYear(periodDefault?.startYear)
        || getCurrentAcademicStartYear();
}

function ensurePeriod(year, makeDefault = false) {
    const startYear = validStartYear(year);
    if (!startYear) return null;
    let period = periods.find((item) => item.startYear === startYear);
    const now = new Date().toISOString();
    if (!period) {
        period = { startYear, label: `${startYear}/${startYear + 1}`, isDefault: false, createdAt: now, updatedAt: now };
        periods.push(period);
    }
    if (makeDefault) setDefaultPeriod(startYear, false);
    periods = periods.sort((a, b) => a.startYear - b.startYear);
    database.periods = periods;
    return period;
}

function setDefaultPeriod(year, updateActive = true) {
    const startYear = validStartYear(year);
    if (!startYear) return;
    ensurePeriod(startYear, false);
    periods.forEach((period) => {
        period.isDefault = period.startYear === startYear;
        if (period.isDefault) period.updatedAt = new Date().toISOString();
    });
    settings.startYear = startYear;
    database.settings = { ...database.settings, ...settings, startYear };
    database.periods = periods;
    if (updateActive) {
        activeStartYear = startYear;
        currentDate = new Date(startYear, 6, 1);
    }
}

function getPeriodEventCount(startYear) {
    const range = getAcademicRangeForYear(startYear);
    return database.events.filter((event) => eventIntersectsRange(event, range.start, range.end)).length;
}

function getPeriodEvents(startYear) {
    const range = getAcademicRangeForYear(startYear);
    return database.events.filter((event) => eventIntersectsRange(event, range.start, range.end));
}

function inferAcademicYear(dateStr) {
    const date = parseDate(dateStr);
    if (Number.isNaN(date.getTime())) return Number(activeStartYear) || Number(settings.startYear);
    return date.getMonth() >= 6 ? date.getFullYear() : date.getFullYear() - 1;
}

function getAcademicRange(scope = "year") {
    const sy = Number(activeStartYear);
    if (scope === "semester1") return { start: `${sy}-07-01`, end: `${sy}-12-31` };
    if (scope === "semester2") return { start: `${sy + 1}-01-01`, end: `${sy + 1}-06-30` };
    return { start: `${sy}-07-01`, end: `${sy + 1}-06-30` };
}

function getAcademicRangeForYear(startYear) {
    const sy = Number(startYear);
    return { start: `${sy}-07-01`, end: `${sy + 1}-06-30` };
}

function getMonthRange(year, month) {
    return {
        start: formatDate(new Date(year, month, 1)),
        end: formatDate(new Date(year, month + 1, 0))
    };
}

function shiftDateByYears(dateStr, delta) {
    const date = parseDate(dateStr);
    if (Number.isNaN(date.getTime())) return dateStr;
    const targetYear = date.getFullYear() + delta;
    const targetMonth = date.getMonth();
    const maxDate = new Date(targetYear, targetMonth + 1, 0).getDate();
    return formatDate(new Date(targetYear, targetMonth, Math.min(date.getDate(), maxDate)));
}

function getRangeEvents(rangeStart, rangeEnd, filtered = true) {
    return database.events
        .filter((event) => eventIntersectsRange(event, rangeStart, rangeEnd))
        .filter((event) => !filtered || activeFilters.has(event.type))
        .sort((a, b) => a.start.localeCompare(b.start) || a.title.localeCompare(b.title));
}

function getYearEvents(filtered = true) {
    const range = getAcademicRange("year");
    return getRangeEvents(range.start, range.end, filtered);
}

function getDayEvents(dateStr) {
    return database.events
        .filter((event) => activeFilters.has(event.type))
        .filter((event) => dateStr >= event.start && dateStr <= event.end)
        .sort((a, b) => a.start.localeCompare(b.start) || a.title.localeCompare(b.title));
}

function eventIntersectsRange(event, rangeStart, rangeEnd) {
    return event.start <= rangeEnd && event.end >= rangeStart;
}

function getAcademicMonths(scope) {
    const sy = Number(activeStartYear);
    const semester1 = [6, 7, 8, 9, 10, 11].map((m) => ({ m, y: sy }));
    const semester2 = [0, 1, 2, 3, 4, 5].map((m) => ({ m, y: sy + 1 }));
    if (scope === "semester1") return semester1;
    if (scope === "semester2") return semester2;
    return [...semester1, ...semester2];
}

function changeMonth(delta) {
    currentDate.setMonth(currentDate.getMonth() + delta);
    saveUiState();
    renderDashboardView();
}

function goToday() {
    currentDate = new Date();
    saveUiState();
    renderDashboardView();
}

function renderDashboardView() {
    const report = $("#export-dashboard");
    $("#month-nav").style.display = dashboardMode === "month" ? "flex" : "none";
    $("#current-month-year").textContent = `${monthNames[currentDate.getMonth()]} ${currentDate.getFullYear()}`;
    const months = dashboardMode === "month"
        ? [{ m: currentDate.getMonth(), y: currentDate.getFullYear() }]
        : getAcademicMonths(dashboardMode === "year" ? "year" : dashboardMode);
    const gridClass = dashboardMode === "year"
        ? "dashboard-grid model-a-year"
        : months.length <= 6 ? "dashboard-grid semester" : "dashboard-grid";

    report.innerHTML = `
        <section class="report-sheet dashboard-sheet print-dashboard-model-a">
            ${renderReportTitle(`KALENDER PENDIDIKAN ${escapeHtml(settings.schoolName)}`, getDashboardSubtitle())}
            <div class="${gridClass}">
                ${months.map((item) => renderMonthCard(item.y, item.m)).join("")}
            </div>
            ${renderDashboardSummary(months)}
            ${renderSignatureBlock("kepala")}
        </section>
    `;
}

function getDashboardSubtitle() {
    const yearText = `TAHUN PELAJARAN ${activeStartYear}/${activeStartYear + 1}`;
    if (dashboardMode === "semester1") return `SEMESTER I - ${yearText}`;
    if (dashboardMode === "semester2") return `SEMESTER II - ${yearText}`;
    if (dashboardMode === "month") return `${monthNames[currentDate.getMonth()].toUpperCase()} ${currentDate.getFullYear()} - ${yearText}`;
    return yearText;
}

function renderMonthCard(year, month) {
    const first = new Date(year, month, 1);
    const start = new Date(year, month, 1 - first.getDay());
    const rows = [];
    for (let week = 0; week < 6; week++) {
        const cells = [];
        for (let day = 0; day < 7; day++) {
            const cellDate = new Date(start);
            cellDate.setDate(start.getDate() + week * 7 + day);
            const dateStr = formatDate(cellDate);
            const other = cellDate.getMonth() !== month;
            const nonWorking = !isWorkingDay(cellDate);
            const events = other ? [] : getDayEvents(dateStr);
            const classes = [
                other ? "other-month" : "",
                nonWorking ? "non-working" : ""
            ].filter(Boolean).join(" ");
            cells.push(`
                <td class="${classes}" data-date="${dateStr}" title="${escapeHtml(events.map((event) => event.title).join(" | "))}">
                    <span class="date-number">${cellDate.getDate()}</span>
                    ${events.slice(0, 3).map((event) => renderEventMini(event)).join("")}
                </td>
            `);
        }
        rows.push(`<tr>${cells.join("")}</tr>`);
    }

    const range = getMonthRange(year, month);
    const monthEvents = getRangeEvents(range.start, range.end, true);
    const eventListHtml = monthEvents.length > 0 ? `
        <div class="month-event-list">
            ${monthEvents.slice(0, 8).map((event) => {
                const category = categories[event.type] || categories.lainnya;
                const d1 = parseDate(event.start).getDate();
                const d2 = parseDate(event.end).getDate();
                const dateLabel = d1 === d2 ? `${d1}` : `${d1}-${d2}`;
                return `
                    <div class="month-event-item">
                        <span class="event-dot" style="background:${category.color}"></span>
                        <span class="event-date">${dateLabel}:</span>
                        <span class="event-text">${escapeHtml(event.title)}</span>
                    </div>
                `;
            }).join("")}
            ${monthEvents.length > 8 ? `<div class="month-event-more">...dan ${monthEvents.length - 8} lainnya</div>` : ""}
        </div>
    ` : "";

    return `
        <section class="dashboard-month-card">
            <table class="month-calendar export-table">
                <thead>
                    <tr class="month-title-row">
                        <th>${month + 1}</th>
                        <th colspan="5">${monthNames[month]}</th>
                        <th>${year}</th>
                    </tr>
                    <tr>${dayNames.map((day, index) => `<th class="${index === 0 || (Number(settings.workDays) === 5 && index === 6) ? "non-working-head" : ""}">${day}</th>`).join("")}</tr>
                </thead>
                <tbody>${rows.join("")}</tbody>
            </table>
            ${eventListHtml}
        </section>
    `;
}

function renderEventMini(event) {
    const category = categories[event.type] || categories.lainnya;
    const light = category.textColor === "black" ? "light-text" : "";
    return `<span class="event-mini ${light}" data-event-id="${event.id}" style="background:${category.color}">${escapeHtml(event.title)}</span>`;
}

function renderDashboardSummary(months) {
    const ranges = months.map((item) => getMonthRange(item.y, item.m));
    const relevant = database.events.filter((event) => ranges.some((range) => eventIntersectsRange(event, range.start, range.end)));
    const nonEffective = relevant.filter((event) => event.isNonEffective).length;
    const effective = relevant.length - nonEffective;
    return `
        <div class="metric-grid no-break no-export">
            <div class="metric-card"><span>Total Kegiatan</span><strong>${relevant.length}</strong></div>
            <div class="metric-card"><span>Tetap KBM</span><strong>${effective}</strong></div>
            <div class="metric-card"><span>Hari Tidak Efektif</span><strong>${nonEffective}</strong></div>
            <div class="metric-card"><span>Hari Kerja</span><strong>${settings.workDays}</strong></div>
        </div>
        <div class="dashboard-export-summary export-only no-break">
            ${buildAcademicSummaryTable("table-dashboard-effective-summary")}
        </div>
    `;
}

function buildDashboardStatsTable(summary) {
    return `
        <table id="table-dashboard-stats" class="summary-table dashboard-stats-table export-table">
            <tbody>
                <tr><td class="text-left">TOTAL KEGIATAN</td><td>${summary.relevant}</td><td>Kegiatan</td></tr>
                <tr><td class="text-left">TETAP KBM</td><td>${summary.effective}</td><td>Kegiatan</td></tr>
                <tr><td class="text-left">HARI TIDAK EFEKTIF</td><td>${summary.nonEffective}</td><td>Kegiatan</td></tr>
                <tr><td class="text-left">HARI KERJA</td><td>${settings.workDays}</td><td>Hari/Minggu</td></tr>
            </tbody>
        </table>
    `;
}

function handleReportClick(event) {
    const eventBadge = event.target.closest("[data-event-id]");
    if (eventBadge) {
        editEvent(eventBadge.dataset.eventId);
        return;
    }
    const actionButton = event.target.closest("[data-row-action]");
    if (actionButton) {
        const id = actionButton.dataset.eventId;
        if (actionButton.dataset.rowAction === "edit") editEvent(id);
        if (actionButton.dataset.rowAction === "delete") deleteEvent(id);
        return;
    }
    const dayCell = event.target.closest("[data-date]");
    if (dayCell && activeView === "dashboard") openEventModal(dayCell.dataset.date);
}

function renderAkademikView() {
    const scope = $("#academic-scope").value;
    const months = getAcademicMonths(scope);
    const pages = [];
    if (academicFormat === "model-b" || academicFormat === "all") pages.push(buildAcademicPage("model-b", months));
    if (academicFormat === "model-c" || academicFormat === "all") pages.push(buildAcademicPage("model-c", months));
    $("#export-akademik").innerHTML = pages.join("");
}

function buildAcademicPage(model, months) {
    const label = model === "model-b" ? "Model B" : "Model C";
    const table = model === "model-b" ? buildModelBTable(months) : buildModelCTable(months);
    const titleLine = model === "model-c"
        ? `KALENDER PENDIDIKAN ${activeStartYear}-${activeStartYear + 1}`
        : `TAHUN PELAJARAN ${activeStartYear}-${activeStartYear + 1}`;
    return `
        <section class="report-sheet akademik-sheet ${model}">
            ${renderReportTitle(`KALENDER PENDIDIKAN ${escapeHtml(settings.schoolName)}`, titleLine)}
            <div class="table-scroll">${table}</div>
            <div class="report-footer-grid">
                <div>
                    ${buildAcademicSummaryTable()}
                    <div class="legend-list no-export">${buildLegend()}</div>
                </div>
                ${renderSignatureBlock("kepala", true)}
            </div>
            <p class="format-note no-export">${label}</p>
        </section>
    `;
}

function buildModelBTable(months) {
    const columnCount = getModelBColumnCount(months);
    const head = `
        <thead>
            <tr>
                <th class="col-no" rowspan="2">NO</th>
                <th class="col-month" rowspan="2">BULAN</th>
                <th colspan="${columnCount}">TANGGAL</th>
            </tr>
            <tr>${Array.from({ length: columnCount }, (_, i) => {
                const dayIndex = i % 7;
                const nonWorking = dayIndex === 0 || (Number(settings.workDays) === 5 && dayIndex === 6);
                return `<th class="date-col day-head ${nonWorking ? "non-working" : ""}"><span>${dayNames[dayIndex]}</span></th>`;
            }).join("")}</tr>
        </thead>
    `;
    const body = months.map((item, index) => {
        const firstDate = new Date(item.y, item.m, 1);
        const lastDate = new Date(item.y, item.m + 1, 0);
        const startOffset = firstDate.getDay();
        const cells = [];
        for (let col = 0; col < columnCount; col++) {
            const date = new Date(item.y, item.m, 1 - startOffset + col);
            if (date > lastDate) {
                cells.push(`<td class="date-cell blank-date"></td>`);
                continue;
            }
            const otherMonth = date < firstDate;
            const dateStr = formatDate(date);
            const dayEvents = otherMonth ? [] : getDayEvents(dateStr);
            const nonWorking = !isWorkingDay(date);
            const style = dayEvents.length ? getMultiEventStyle(dayEvents) : "";
            const light = dayEvents.length === 1 && (categories[dayEvents[0].type] || {}).textColor === "black";
            const classes = [
                "date-cell",
                otherMonth ? "other-month" : "",
                nonWorking ? "non-working" : "",
                dayEvents.length ? "event-cell" : "",
                light ? "light-text" : ""
            ].filter(Boolean).join(" ");
            cells.push(`
                <td class="${classes}" style="${style}" title="${escapeHtml(dayEvents.map((event) => event.title).join(" | "))}">
                    ${date.getDate()}
                </td>
            `);
        }
        return `<tr><td>${index + 1}</td><td class="col-month">${monthNames[item.m]} ${item.y}</td>${cells.join("")}</tr>`;
    }).join("");
    return `<table id="table-akademik-model-b" class="academic-board model-b-table export-table">${head}<tbody>${body}</tbody></table>`;
}

function getModelBColumnCount(months) {
    return Math.max(...months.map((item) => {
        const firstDate = new Date(item.y, item.m, 1);
        const daysInMonth = new Date(item.y, item.m + 1, 0).getDate();
        return firstDate.getDay() + daysInMonth;
    }), 31);
}

function buildModelCTable(months) {
    const head = `
        <thead>
            <tr>
                <th class="col-no" rowspan="2">NO</th>
                <th class="col-month" rowspan="2">BULAN</th>
                <th colspan="31">TANGGAL</th>
            </tr>
            <tr>${Array.from({ length: 31 }, (_, i) => `<th class="date-col">${i + 1}</th>`).join("")}</tr>
        </thead>
    `;
    const body = months.map((item, index) => {
        const daysInMonth = new Date(item.y, item.m + 1, 0).getDate();
        const cells = [];
        for (let day = 1; day <= 31; day++) {
            if (day > daysInMonth) {
                cells.push(`<td class="date-cell invalid-date"></td>`);
                continue;
            }
            const date = new Date(item.y, item.m, day);
            const dateStr = formatDate(date);
            const dayEvents = getDayEvents(dateStr);
            const nonWorking = !isWorkingDay(date);
            const style = dayEvents.length ? getMultiEventStyle(dayEvents) : "";
            const light = dayEvents.length === 1 && (categories[dayEvents[0].type] || {}).textColor === "black";
            const classes = [
                "date-cell",
                nonWorking ? "non-working" : "",
                dayEvents.length ? "event-cell" : "",
                light ? "light-text" : ""
            ].filter(Boolean).join(" ");
            cells.push(`<td class="${classes}" style="${style}" title="${escapeHtml(dayEvents.map((event) => event.title).join(" | "))}">${dayNames[date.getDay()]}</td>`);
        }
        return `<tr><td>${index + 1}</td><td class="col-month">${monthNames[item.m]} ${item.y}</td>${cells.join("")}</tr>`;
    }).join("");
    return `<table id="table-akademik-model-c" class="academic-board model-c-table export-table">${head}<tbody>${body}</tbody></table>`;
}

function getMultiEventStyle(events) {
    if (!events.length) return "";
    const colors = events.slice(0, 3).map((event) => (categories[event.type] || categories.lainnya).color);
    if (colors.length === 1) return `background:${colors[0]}`;
    if (colors.length === 2) return `background:linear-gradient(135deg, ${colors[0]} 0%, ${colors[0]} 50%, ${colors[1]} 50%, ${colors[1]} 100%)`;
    return `background:linear-gradient(135deg, ${colors[0]} 0%, ${colors[0]} 33.33%, ${colors[1]} 33.33%, ${colors[1]} 66.66%, ${colors[2]} 66.66%, ${colors[2]} 100%)`;
}

function buildAcademicSummaryTable(tableId = "table-akademik-summary") {
    const rows = getAcademicSummaryRowsData();
    return `
        <table id="${tableId}" class="summary-table export-table">
            <thead><tr><th colspan="5"></th><th colspan="3">Semester I</th><th colspan="3">Semester II</th></tr></thead>
            <tbody>
                ${rows.map((row) => `<tr><td class="text-left" colspan="5">${row.label}</td><td colspan="3">${row.semester1}</td><td colspan="3">${row.semester2}</td></tr>`).join("")}
            </tbody>
        </table>
    `;
}

function getAcademicSummaryRowsData() {
    const semester1 = hitungSemester(getAcademicMonths("semester1"));
    const semester2 = hitungSemester(getAcademicMonths("semester2"));
    return [
        { label: "Hari Efektif KBM", semester1: `${semester1.total.hariEfektif} Hari`, semester2: `${semester2.total.hariEfektif} Hari` },
        { label: "Pekan Efektif KBM", semester1: `${semester1.total.pekanEfektif} Pekan`, semester2: `${semester2.total.pekanEfektif} Pekan` },
        { label: "Pekan Tidak Efektif KBM", semester1: `${semester1.total.pekanTidakEfektif} Pekan`, semester2: `${semester2.total.pekanTidakEfektif} Pekan` }
    ];
}

function buildLegend() {
    const items = [
        `<span class="legend-item"><span class="swatch" style="background:#fee2e2"></span> Hari tidak kerja</span>`,
        ...Object.entries(categories)
            .filter(([key]) => activeFilters.has(key))
            .map(([, category]) => `<span class="legend-item"><span class="swatch" style="background:${category.color}"></span> ${escapeHtml(category.label)}</span>`)
    ];
    return items.join("");
}

function renderEfektifView() {
    const semester1 = hitungSemester(getAcademicMonths("semester1"));
    const semester2 = hitungSemester(getAcademicMonths("semester2"));
    $("#export-efektif").innerHTML = [
        buildEffectiveSheet("Ganjil", semester1, "#93c5fd", `${activeStartYear}-07-13`),
        buildEffectiveSheet("Genap", semester2, "#fef08a", `${activeStartYear + 1}-01-02`)
    ].join("");
}

function buildEffectiveSheet(semesterName, result, accent, defaultDate) {
    const total = result.total;
    const ttdDate = semesterName === "Genap" ? formatDateIndo(defaultDate) : formatDateIndo(settings.ttdDate);
    return `
        <section class="report-sheet effective-sheet" style="--effective-accent:${accent}">
            <div class="effective-title">ANALISIS PEKAN EFEKTIF</div>
            <table class="effective-identity-table export-table">
                <tbody>
                    <tr>
                        <td>Nama Sekolah</td><td>: ${escapeHtml(settings.schoolName)}</td>
                        <td>Kelas</td><td>: ${escapeHtml(settings.className)}</td>
                    </tr>
                    <tr>
                        <td>Mata pelajaran</td><td>: ${escapeHtml(settings.subjectName)}</td>
                        <td>Tahun Ajaran</td><td>: ${activeStartYear}-${activeStartYear + 1}</td>
                    </tr>
                    <tr>
                        <td>Nama Guru</td><td>: ${escapeHtml(settings.teacherName)}</td>
                        <td>Semester</td><td>: ${semesterName}</td>
                    </tr>
                </tbody>
            </table>
            <h3 class="effective-section-title">A. Pekan Efektif</h3>
            ${renderEffectiveTable(result)}
            <h3 class="effective-section-title">B. Jumlah Jam Tatap Muka</h3>
            <table class="effective-summary export-table">
                <tbody>
                    <tr><td>a.</td><td>Jumlah pekan efektif KBM</td><td>:</td><td>${total.pekanEfektif}</td><td>Pekan</td></tr>
                    <tr><td>b.</td><td>Pekan tidak efektif KBM</td><td>:</td><td>${total.pekanTidakEfektif}</td><td>Pekan</td></tr>
                    <tr><td>c.</td><td>Jumlah hari efektif KBM</td><td>:</td><td>${total.hariEfektif}</td><td>Hari</td></tr>
                    <tr><td>d.</td><td>Alokasi jam per minggu</td><td>:</td><td>${settings.jamPerMinggu}</td><td>Jam</td></tr>
                    <tr><td>e.</td><td>Jumlah jam Efektif</td><td>:</td><td>${total.jamEfektif}</td><td>Jam</td></tr>
                </tbody>
            </table>
            ${renderSignatureBlock("guru", false, ttdDate)}
        </section>
    `;
}

function renderEffectiveTable(result) {
    return `
        <table class="effective-board export-table">
            <thead>
                <tr>
                    <th rowspan="2">No</th>
                    <th rowspan="2">Bulan</th>
                    <th colspan="2">Tersedia</th>
                    <th colspan="2">Waktu Tidak Efektif</th>
                    <th colspan="2">Waktu Efektif</th>
                </tr>
                <tr>
                    <th>Pekan</th>
                    <th>Hari</th>
                    <th>Libur Akhir Pekan</th>
                    <th>Libur Hari Lainnya</th>
                    <th>Hari</th>
                    <th>Pekan</th>
                </tr>
            </thead>
            <tbody>
                ${result.data.map((row) => `
                    <tr>
                        <td>${row.no}</td>
                        <td class="text-left">${escapeHtml(row.bulan)}</td>
                        <td>${row.pekanTersedia}</td>
                        <td>${row.hariDalamBulan}</td>
                        <td>${row.hariNonKerja}</td>
                        <td>${row.hariTidakEfektif}</td>
                        <td>${row.hariEfektif}</td>
                        <td>${row.pekanEfektif}</td>
                    </tr>
                `).join("")}
            </tbody>
            <tfoot>
                <tr>
                    <th></th><th></th>
                    <th>${result.total.pekanTersedia}</th>
                    <th>${result.total.hariDalamBulan}</th>
                    <th>${result.total.hariNonKerja}</th>
                    <th>${result.total.hariTidakEfektif}</th>
                    <th>${result.total.hariEfektif}</th>
                    <th>${result.total.pekanEfektif}</th>
                </tr>
            </tfoot>
        </table>
    `;
}

function hitungSemester(months) {
    const data = [];
    const total = {
        pekanTersedia: 0,
        hariDalamBulan: 0,
        hariNonKerja: 0,
        hariTidakEfektif: 0,
        pekanTidakEfektif: 0,
        pekanEfektif: 0,
        hariEfektif: 0,
        jamEfektif: 0
    };

    months.forEach((item, index) => {
        const daysInMonth = new Date(item.y, item.m + 1, 0).getDate();
        let hariNonKerja = 0;
        let hariTidakEfektif = 0;
        for (let day = 1; day <= daysInMonth; day++) {
            const date = new Date(item.y, item.m, day);
            const dateStr = formatDate(date);
            if (!isWorkingDay(date)) {
                hariNonKerja++;
                continue;
            }
            const dayEvents = getDayEvents(dateStr).filter((event) => event.isNonEffective);
            if (dayEvents.length) hariTidakEfektif++;
        }
        const hariEfektif = Math.max(0, daysInMonth - hariNonKerja - hariTidakEfektif);
        const pekanTersedia = Math.ceil(daysInMonth / 7);
        const pekanEfektif = effectiveWeeksFromDays(hariEfektif);
        const pekanTidakEfektif = Math.max(0, pekanTersedia - pekanEfektif);
        const jamEfektif = pekanEfektif * settings.jamPerMinggu;
        const row = {
            no: index + 1,
            bulan: `${monthNames[item.m]} ${item.y}`,
            pekanTersedia,
            hariDalamBulan: daysInMonth,
            hariNonKerja,
            hariTidakEfektif,
            pekanTidakEfektif,
            pekanEfektif,
            hariEfektif,
            jamEfektif
        };
        data.push(row);
        Object.keys(total).forEach((key) => { total[key] += row[key] || 0; });
    });
    return { data, total };
}

function effectiveWeeksFromDays(days) {
    const workDays = Number(settings.workDays) || 5;
    const full = Math.floor(days / workDays);
    const remainder = days % workDays;
    return full + (remainder >= Math.ceil(workDays / 2) ? 1 : 0);
}

function renderRekapView() {
    const filtered = filterRekapEvents();
    const effective = filtered.filter((event) => !event.isNonEffective);
    const nonEffective = filtered.filter((event) => event.isNonEffective);
    renderRekapRows("#table-rekap-effective tbody", effective, "Tetap KBM");
    renderRekapRows("#table-rekap-non-effective tbody", nonEffective, "Hari Tidak Efektif KBM");
}

function filterRekapEvents() {
    const semester = $("#filter-rekap-semester").value;
    const monthValue = $("#filter-rekap-bulan").value;
    let range = getAcademicRange(semester === "semester1" ? "semester1" : semester === "semester2" ? "semester2" : "year");
    if (monthValue !== "all") {
        const [year, month] = monthValue.split("-").map(Number);
        range = getMonthRange(year, month - 1);
    }
    return getRangeEvents(range.start, range.end, true);
}

function renderRekapRows(selector, rows, status) {
    const tbody = $(selector);
    if (!rows.length) {
        tbody.innerHTML = `<tr><td colspan="8" class="empty-row">Tidak ada data.</td></tr>`;
        return;
    }
    tbody.innerHTML = rows.map((event, index) => {
        const category = categories[event.type] || categories.lainnya;
        const statusClass = event.isNonEffective ? "status-non-effective" : "status-effective";
        const light = category.textColor === "black" ? "light-text" : "";
        return `
            <tr>
                <td>${index + 1}</td>
                <td class="text-left"><strong>${escapeHtml(event.title)}</strong>${event.notes ? `<br><small>${escapeHtml(event.notes)}</small>` : ""}</td>
                <td>${formatDateIndo(event.start)}</td>
                <td>${formatDateIndo(event.end)}</td>
                <td>${calculateDays(event.start, event.end)} Hari</td>
                <td><span class="category-pill ${light}" style="background:${category.color}">${escapeHtml(category.label)}</span></td>
                <td class="${statusClass}">${status}</td>
                <td class="no-export">
                    <button type="button" class="event-action" data-row-action="edit" data-event-id="${event.id}" title="Edit"><i class="fa-solid fa-pen"></i></button>
                </td>
            </tr>
        `;
    }).join("");
}

async function openCopyPeriodDialog() {
    const currentRange = getAcademicRangeForYear(activeStartYear);
    const sourceEvents = database.events.filter((event) => eventIntersectsRange(event, currentRange.start, currentRange.end));
    if (!sourceEvents.length) {
        toast("Tidak ada kegiatan pada periode aktif untuk disalin.", "info");
        return;
    }
    const result = await Swal.fire({
        title: "Salin Kegiatan Periode",
        width: "560px",
        html: `
            <div class="copy-period-dialog">
                <p>Salin ${sourceEvents.length} kegiatan dari tahun pelajaran <strong>${activeStartYear}/${activeStartYear + 1}</strong> ke periode tujuan.</p>
                <label>Periode Tujuan
                    <select id="copy-period-target">
                        <option value="1">${activeStartYear + 1}/${activeStartYear + 2}</option>
                        <option value="-1">${activeStartYear - 1}/${activeStartYear}</option>
                    </select>
                </label>
                <label class="checkbox-row">
                    <input type="checkbox" id="copy-period-replace">
                    <span>Hapus kegiatan yang sudah ada di periode tujuan sebelum menyalin</span>
                </label>
                <p class="dialog-hint">Tanggal kegiatan akan digeser satu tahun sesuai tujuan, sedangkan kategori, status KBM, dan catatan tetap dipertahankan.</p>
            </div>
        `,
        showCancelButton: true,
        confirmButtonText: "Salin Kegiatan",
        cancelButtonText: "Batal",
        focusConfirm: false,
        preConfirm: () => ({
            delta: Number(document.getElementById("copy-period-target").value),
            replace: document.getElementById("copy-period-replace").checked
        })
    });
    if (!result.isConfirmed) return;
    const copied = copyPeriodEvents(result.value.delta, result.value.replace);
    if (!copied) {
        toast("Tidak ada kegiatan baru yang disalin.", "info");
        return;
    }
    setupYearSelectors();
    saveUiState();
    markDataChanged();
    renderRekapView();
    toast(`${copied} kegiatan berhasil disalin`);
}

function copyPeriodEvents(delta, replaceTarget = false) {
    const currentRange = getAcademicRangeForYear(activeStartYear);
    const targetRange = getAcademicRangeForYear(activeStartYear + delta);
    const sourceEvents = database.events.filter((event) => eventIntersectsRange(event, currentRange.start, currentRange.end));
    if (!sourceEvents.length) return 0;
    if (replaceTarget) {
        database.events = database.events.filter((event) => !eventIntersectsRange(event, targetRange.start, targetRange.end));
    }
    const existingKeys = new Set(database.events.map((event) => [
        event.title.toLowerCase(),
        event.type,
        event.start,
        event.end
    ].join("|")));
    const copiedEvents = sourceEvents
        .map((event) => {
            const start = shiftDateByYears(event.start, delta);
            const end = shiftDateByYears(event.end, delta);
            return {
                ...event,
                id: createId(),
                academicYear: inferAcademicYear(start),
                start,
                end
            };
        })
        .filter((event) => {
            const key = [event.title.toLowerCase(), event.type, event.start, event.end].join("|");
            if (existingKeys.has(key)) return false;
            existingKeys.add(key);
            return true;
        });
    database.events.push(...copiedEvents);
    return copiedEvents.length;
}

function openEventModal(dateStr = "") {
    $("#event-form").reset();
    $("#event-id").value = "";
    $("#modal-title").textContent = "Tambah Kegiatan";
    $("#btn-delete-event").classList.add("hidden");
    const selectedType = Object.keys(categories)[0] || "lainnya";
    $("#event-type").value = selectedType;
    $("#event-isNonEffective").checked = Boolean((categories[selectedType] || {}).defaultNonEffective);
    if (dateStr) {
        $("#event-start").value = dateStr;
        $("#event-end").value = dateStr;
        $("#event-end").min = dateStr;
    }
    $("#event-modal").classList.add("open");
    $("#event-modal").setAttribute("aria-hidden", "false");
}

function closeEventModal() {
    $("#event-modal").classList.remove("open");
    $("#event-modal").setAttribute("aria-hidden", "true");
}

function editEvent(id) {
    const event = database.events.find((item) => item.id === id);
    if (!event) return;
    $("#event-id").value = event.id;
    $("#event-title").value = event.title;
    $("#event-start").value = event.start;
    $("#event-end").value = event.end;
    $("#event-end").min = event.start;
    $("#event-type").value = event.type;
    $("#event-notes").value = event.notes || "";
    $("#event-isNonEffective").checked = Boolean(event.isNonEffective);
    $("#modal-title").textContent = "Edit Kegiatan";
    $("#btn-delete-event").classList.remove("hidden");
    $("#event-modal").classList.add("open");
    $("#event-modal").setAttribute("aria-hidden", "false");
}

function handleEventSubmit(event) {
    event.preventDefault();
    const id = $("#event-id").value || createId();
    const start = $("#event-start").value;
    const end = $("#event-end").value;
    if (end < start) {
        toast("Tanggal berakhir tidak boleh lebih awal.", "error");
        return;
    }
    const record = {
        id,
        academicYear: inferAcademicYear(start),
        title: $("#event-title").value.trim(),
        start,
        end,
        type: $("#event-type").value,
        isNonEffective: $("#event-isNonEffective").checked,
        notes: $("#event-notes").value.trim()
    };
    const index = database.events.findIndex((item) => item.id === id);
    if (index >= 0) database.events[index] = record;
    else database.events.push(record);
    setupYearSelectors();
    applySettingsToDOM();
    closeEventModal();
    saveUiState();
    markDataChanged();
    renderCurrentView();
    toast(index >= 0 ? "Kegiatan diperbarui" : "Kegiatan ditambahkan");
}

async function deleteEvent(id) {
    if (!id) return;
    const confirmed = await confirmDialog("Hapus kegiatan ini?", "Data yang dihapus tidak dapat dikembalikan.");
    if (!confirmed) return;
    database.events = database.events.filter((event) => event.id !== id);
    closeEventModal();
    markDataChanged();
    renderCurrentView();
    toast("Kegiatan dihapus");
}

function resetCategoryForm() {
    $("#category-form").reset();
    $("#category-original-key").value = "";
    $("#category-key").disabled = false;
    $("#category-color").value = "#3b82f6";
    $("#category-text-color").value = "white";
    $("#category-default-non-effective").checked = false;
}

function saveCategory(event) {
    event.preventDefault();
    const originalKey = $("#category-original-key").value;
    const key = sanitizeCategoryKey($("#category-key").value);
    if (!key) {
        toast("Kode kategori wajib diisi.", "error");
        return;
    }
    if (!originalKey && categories[key]) {
        toast("Kode kategori sudah digunakan.", "error");
        return;
    }
    if (originalKey && originalKey !== key && categories[key]) {
        toast("Kode kategori pengganti sudah digunakan.", "error");
        return;
    }
    const category = {
        label: $("#category-label").value.trim(),
        color: normalizeColor($("#category-color").value),
        textColor: $("#category-text-color").value === "black" ? "black" : "white",
        defaultNonEffective: $("#category-default-non-effective").checked
    };
    if (!category.label) {
        toast("Nama kategori wajib diisi.", "error");
        return;
    }
    if (originalKey && originalKey !== key) {
        database.events.forEach((item) => {
            if (item.type === originalKey) item.type = key;
        });
        delete categories[originalKey];
        activeFilters.delete(originalKey);
    }
    categories[key] = category;
    activeFilters.add(key);
    database.categories = categories;
    setupFilters();
    setupCategoryOptions();
    renderCategoryTable();
    saveUiState();
    markDataChanged();
    renderCurrentView();
    toast("Kategori disimpan");
}

function renderCategoryTable() {
    const tbody = $("#table-categories tbody");
    if (!tbody) return;
    tbody.innerHTML = Object.entries(categories).map(([key, category]) => {
        const usedCount = database.events.filter((event) => event.type === key).length;
        const textClass = category.textColor === "black" ? "light-text" : "";
        return `
            <tr>
                <td><code>${escapeHtml(key)}</code></td>
                <td><span class="category-pill ${textClass}" style="background:${category.color}">${escapeHtml(category.label)}</span></td>
                <td class="text-left">${escapeHtml(category.label)}<br><small>${usedCount} kegiatan</small></td>
                <td>${category.defaultNonEffective ? "Hari Tidak Efektif" : "Tetap KBM"}</td>
                <td>
                    <button type="button" class="event-action" data-category-action="edit" data-key="${key}" title="Edit"><i class="fa-solid fa-pen"></i></button>
                    <button type="button" class="event-action" data-category-action="delete" data-key="${key}" title="Hapus"><i class="fa-solid fa-trash"></i></button>
                </td>
            </tr>
        `;
    }).join("");
}

function renderPeriodTable() {
    const tbody = $("#table-periods tbody");
    if (!tbody) return;
    if (!periods.length) {
        tbody.innerHTML = `<tr><td colspan="4" class="empty-row">Belum ada periode pembelajaran.</td></tr>`;
        return;
    }
    tbody.innerHTML = periods.map((period) => {
        const count = getPeriodEventCount(period.startYear);
        return `
            <tr>
                <td><strong>${period.startYear}/${period.startYear + 1}</strong></td>
                <td>${count} kegiatan</td>
                <td>${period.isDefault ? '<span class="state-pill ready">Default</span>' : '<span class="state-pill">Aktif bila dipilih</span>'}</td>
                <td>
                    <button type="button" class="event-action" data-period-action="default" data-year="${period.startYear}" title="Jadikan default"><i class="fa-solid fa-star"></i></button>
                    <button type="button" class="event-action" data-period-action="copy-prev" data-year="${period.startYear}" title="Salin dari periode sebelumnya"><i class="fa-solid fa-arrow-left"></i></button>
                    <button type="button" class="event-action" data-period-action="copy-next" data-year="${period.startYear}" title="Salin dari periode setelahnya"><i class="fa-solid fa-arrow-right"></i></button>
                    <button type="button" class="event-action" data-period-action="delete" data-year="${period.startYear}" title="Hapus periode"><i class="fa-solid fa-trash"></i></button>
                </td>
            </tr>
        `;
    }).join("");
}

async function handlePeriodTableClick(event) {
    const button = event.target.closest("[data-period-action]");
    if (!button) return;
    const year = validStartYear(button.dataset.year);
    if (!year) return;
    const action = button.dataset.periodAction;
    if (action === "default") setPeriodAsDefault(year);
    if (action === "copy-prev") copyEventsIntoPeriod(year, year - 1);
    if (action === "copy-next") copyEventsIntoPeriod(year, year + 1);
    if (action === "delete") deleteLearningPeriod(year);
}

async function addLearningPeriod() {
    const nextYear = Math.max(...periods.map((period) => period.startYear), settings.startYear) + 1;
    const result = await Swal.fire({
        title: "Tambah Periode Pembelajaran",
        input: "number",
        inputLabel: "Tahun Pelajaran Mulai",
        inputValue: nextYear,
        inputAttributes: { min: "2000", max: "2100", step: "1" },
        showCancelButton: true,
        confirmButtonText: "Tambah",
        cancelButtonText: "Batal",
        inputValidator: (value) => {
            const year = validStartYear(value);
            if (!year) return "Tahun pelajaran tidak valid.";
            if (periods.some((period) => period.startYear === year)) return "Periode sudah ada.";
            return undefined;
        }
    });
    if (!result.isConfirmed) return;
    ensurePeriod(Number(result.value));
    renderPeriodTable();
    setupYearSelectors();
    saveUiState();
    markDataChanged();
    toast("Periode pembelajaran ditambahkan");
}

function setPeriodAsDefault(year) {
    setDefaultPeriod(year, true);
    setupYearSelectors();
    populateSettingsForm();
    applySettingsToDOM();
    populateRekapMonthFilter();
    renderPeriodTable();
    saveUiState();
    markDataChanged();
    renderCurrentView();
    toast(`Periode default diubah ke ${year}/${year + 1}`);
}

async function deleteLearningPeriod(year) {
    const count = getPeriodEventCount(year);
    const isDefault = settings.startYear === year;
    if (periods.length <= 1) {
        toast("Minimal harus ada satu periode pembelajaran.", "error");
        return;
    }
    const confirmed = await confirmDialog(
        `Hapus periode ${year}/${year + 1}?`,
        `Sebanyak ${count} kegiatan pada periode ini akan ikut dihapus.`
    );
    if (!confirmed) return;
    const range = getAcademicRangeForYear(year);
    database.events = database.events.filter((event) => !eventIntersectsRange(event, range.start, range.end));
    periods = periods.filter((period) => period.startYear !== year);
    if (isDefault) setDefaultPeriod(periods[0].startYear, true);
    database.periods = periods;
    setupYearSelectors();
    populateSettingsForm();
    applySettingsToDOM();
    populateRekapMonthFilter();
    renderPeriodTable();
    saveUiState();
    markDataChanged();
    renderCurrentView();
    toast("Periode pembelajaran dihapus");
}

function copyEventsIntoPeriod(targetYear, sourceYear) {
    ensurePeriod(sourceYear);
    ensurePeriod(targetYear);
    const sourceEvents = getPeriodEvents(sourceYear);
    if (!sourceEvents.length) {
        toast(`Tidak ada kegiatan di periode ${sourceYear}/${sourceYear + 1}.`, "info");
        return;
    }
    const delta = targetYear - sourceYear;
    const existingKeys = new Set(database.events.map((event) => [
        event.title.toLowerCase(),
        event.type,
        event.start,
        event.end
    ].join("|")));
    const copied = sourceEvents
        .map((event) => {
            const start = shiftDateByYears(event.start, delta);
            const end = shiftDateByYears(event.end, delta);
            return { ...event, id: createId(), academicYear: inferAcademicYear(start), start, end };
        })
        .filter((event) => {
            const key = [event.title.toLowerCase(), event.type, event.start, event.end].join("|");
            if (existingKeys.has(key)) return false;
            existingKeys.add(key);
            return true;
        });
    database.events.push(...copied);
    setupYearSelectors();
    renderPeriodTable();
    saveUiState();
    markDataChanged();
    renderCurrentView();
    toast(`${copied.length} kegiatan disalin ke ${targetYear}/${targetYear + 1}`);
}

async function handleCategoryTableClick(event) {
    const button = event.target.closest("[data-category-action]");
    if (!button) return;
    const key = button.dataset.key;
    if (button.dataset.categoryAction === "edit") editCategory(key);
    if (button.dataset.categoryAction === "delete") deleteCategory(key);
}

function editCategory(key) {
    const category = categories[key];
    if (!category) return;
    $("#category-original-key").value = key;
    $("#category-key").value = key;
    $("#category-key").disabled = false;
    $("#category-label").value = category.label;
    $("#category-color").value = normalizeColor(category.color);
    $("#category-text-color").value = category.textColor;
    $("#category-default-non-effective").checked = Boolean(category.defaultNonEffective);
}

async function deleteCategory(key) {
    if (!categories[key]) return;
    if (Object.keys(categories).length <= 1) {
        toast("Minimal harus ada satu kategori.", "error");
        return;
    }
    const used = database.events.filter((event) => event.type === key);
    if (used.length) {
        const options = Object.fromEntries(Object.entries(categories)
            .filter(([candidate]) => candidate !== key)
            .map(([candidate, category]) => [candidate, category.label]));
        const result = await Swal.fire({
            title: "Kategori masih digunakan",
            text: `${used.length} kegiatan memakai kategori ini. Pilih kategori pengganti sebelum dihapus.`,
            input: "select",
            inputOptions: options,
            inputPlaceholder: "Pilih kategori pengganti",
            showCancelButton: true,
            confirmButtonText: "Pindahkan dan hapus",
            cancelButtonText: "Batal",
            inputValidator: (value) => !value ? "Kategori pengganti wajib dipilih." : undefined
        });
        if (!result.isConfirmed || !result.value) return;
        database.events.forEach((event) => {
            if (event.type === key) event.type = result.value;
        });
    } else {
        const confirmed = await confirmDialog("Hapus kategori ini?", "Kategori akan dihapus dari database.");
        if (!confirmed) return;
    }
    delete categories[key];
    activeFilters.delete(key);
    database.categories = categories;
    setupFilters();
    setupCategoryOptions();
    renderCategoryTable();
    resetCategoryForm();
    saveUiState();
    markDataChanged();
    renderCurrentView();
    toast("Kategori dihapus");
}

async function loadBundledDatabase() {
    try {
        const response = await fetch("kalender_database.xlsx", { cache: "no-store" });
        if (!response.ok) throw new Error("Database tidak ditemukan");
        const buffer = await response.arrayBuffer();
        // Return result of loadWorkbookBuffer
        return loadWorkbookBuffer(buffer, "kalender_database.xlsx", false);
    } catch (error) {
        console.warn("Bundled database not found or failed to load:", error);
        databaseMissingOnBoot = true;
        updateDatabaseStatus("Database Excel tidak ditemukan", "dirty");
        return false;
    }
}

async function recoverDatabaseHandle() {
    if (!window.indexedDB) return;
    try {
        dbFileHandle = await getStoredValue(DATABASE_HANDLE_KEY);
        if (dbFileHandle) dbFileName = dbFileHandle.name || dbFileName;
    } catch (error) {
        dbFileHandle = null;
    }
}

async function loadDatabaseFromHandle(showMessage = true) {
    if (!dbFileHandle || !(await verifyPermission(dbFileHandle, false, false))) return false;
    try {
        const file = await dbFileHandle.getFile();
        const buffer = await file.arrayBuffer();
        loadWorkbookBuffer(buffer, file.name, false);
        dbFileName = file.name;
        if (showMessage) toast("Database Excel dimuat");
        return true;
    } catch (error) {
        console.warn(error);
        return false;
    }
}

async function openDatabase() {
    try {
        if (window.showOpenFilePicker) {
            const [handle] = await window.showOpenFilePicker({
                types: [{ description: "Excel Database", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }],
                multiple: false
            });
            const file = await handle.getFile();
            const buffer = await file.arrayBuffer();

            const success = loadWorkbookBuffer(buffer, file.name);
            if (success) {
                dbFileHandle = handle;
                await storeDatabaseHandle(handle);
                await verifyPermission(handle, true, true);
                toast("Database Excel dimuat");
            }
            return;
        }
    } catch (error) {
        if (error && error.name === "AbortError") return;
        console.warn(error);
    }
    $("#db-file-input").click();
}

async function loadDatabaseFromInput(event) {
    const file = event.target.files[0];
    if (!file) return;
    const buffer = await file.arrayBuffer();
    dbFileHandle = null;
    await storeDatabaseHandle(null);
    loadWorkbookBuffer(buffer, file.name);
    event.target.value = "";
    toast("Database Excel dimuat");
}

function loadWorkbookBuffer(buffer, fileName, rerender = true) {
    const savedState = loadUiState();
    let workbook;
    try {
        workbook = XLSX.read(buffer, { type: "array" });
    } catch (e) {
        console.error("Gagal membaca file Excel:", e);
        toast("File yang dipilih bukan file Excel yang valid atau rusak.", "error");
        return false;
    }

    const parsedData = workbookToDatabase(workbook);

    // Safety check: if parsed settings are empty or invalid, don't overwrite current database
    if (!parsedData.settings || !parsedData.events) {
        console.error("Format database tidak valid.");
        toast("Format file Excel tidak sesuai dengan standar aplikasi.", "error");
        return false;
    }

    database = parsedData;
    settings = normalizeSettings(database.settings);
    categories = normalizeCategories(database.categories);
    database.categories = categories;
    database.events = normalizeEvents(database.events);
    database.periods = normalizePeriods(database.periods, database.events, settings.startYear);
    periods = database.periods;
    settings.startYear = getDefaultPeriodYear();
    database.settings = { ...settings };
    dbFileName = fileName || "kalender_database.xlsx";
    activeFilters = new Set(Object.keys(categories));
    activeStartYear = resolveActiveStartYear(savedState);
    persistActiveYear(activeStartYear);
    currentDate = normalizeSavedDate(savedState.currentDate);
    if (!currentDate || !dateInAcademicYear(currentDate, activeStartYear)) currentDate = new Date(activeStartYear, 6, 1);
    saveLocalDataSnapshot();
    databaseMissingOnBoot = false;
    if (rerender) {
        setupYearSelectors();
        setupFilters();
        setupCategoryOptions();
        populateSettingsForm();
        applySettingsToDOM();
        populateRekapMonthFilter();
        renderCategoryTable();
        renderPeriodTable();
        setDatabaseDirty(false);
        renderCurrentView();
        saveUiState();
    }
    return true;
}

function workbookToDatabase(workbook) {
    const settingsSheet = workbook.Sheets.Settings || workbook.Sheets.Pengaturan;
    const eventSheet = workbook.Sheets.Events || workbook.Sheets.Kegiatan;
    const categorySheet = workbook.Sheets.Categories || workbook.Sheets.Kategori;
    const periodSheet = workbook.Sheets.Periods || workbook.Sheets.Periode;
    const parsedSettings = { ...DEFAULT_SETTINGS };
    if (settingsSheet) {
        const rows = XLSX.utils.sheet_to_json(settingsSheet, { header: 1, defval: "" });
        rows.slice(1).forEach((row) => {
            if (!row[0]) return;
            parsedSettings[row[0]] = row[1];
        });
    }
    let parsedCategories = cloneCategories(DEFAULT_CATEGORIES);
    if (categorySheet) {
        const rows = XLSX.utils.sheet_to_json(categorySheet, { defval: "" });
        parsedCategories = {};
        rows.forEach((row) => {
            const key = sanitizeCategoryKey(row.key || row.kode || row.Kode);
            if (!key) return;
            const defaultStatus = row.defaultNonEffective ?? row.defaultTidakEfektif ?? row["Default Tidak Efektif"];
            const parsed = {
                label: row.label || row.nama || row.Nama || key,
                color: row.color || row.warna || row.Warna || "#3b82f6",
                textColor: row.textColor || row.warnaTeks || row["Warna Teks"] || "white"
            };
            if (defaultStatus !== undefined && defaultStatus !== "") parsed.defaultNonEffective = defaultStatus;
            parsedCategories[key] = parsed;
        });
    }
    let parsedEvents = [];
    if (eventSheet) {
        parsedEvents = XLSX.utils.sheet_to_json(eventSheet, { defval: "" }).map((row) => ({
            id: String(row.id || row.ID || createId()),
            academicYear: Number(row.academicYear || row.tahunPelajaran || row.Tahun || parsedSettings.startYear),
            title: String(row.title || row.kegiatan || row.Kegiatan || "").trim(),
            start: normalizeDateCell(row.start || row.tanggalMulai || row["Tanggal Mulai"]),
            end: normalizeDateCell(row.end || row.tanggalBerakhir || row["Tanggal Berakhir"]),
            type: sanitizeCategoryKey(row.type || row.kategori || row.Kategori || "lainnya"),
            isNonEffective: normalizeBoolean(row.isNonEffective ?? row.tidakEfektif ?? row["Tidak Efektif"]),
            notes: String(row.notes || row.catatan || row.Catatan || "")
        })).filter((event) => event.title && event.start && event.end);
    }
    let parsedPeriods = [];
    if (periodSheet) {
        parsedPeriods = XLSX.utils.sheet_to_json(periodSheet, { defval: "" }).map((row) => ({
            startYear: Number(row.startYear || row.tahunMulai || row.Tahun || row.year),
            label: row.label || row.nama || row.Nama,
            isDefault: normalizeBoolean(row.isDefault ?? row.default ?? row.Default),
            createdAt: row.createdAt || row.dibuat || "",
            updatedAt: row.updatedAt || row.diubah || ""
        })).filter((period) => validStartYear(period.startYear));
    }
    const normalizedSettings = normalizeSettings(parsedSettings);
    return {
        settings: normalizedSettings,
        events: parsedEvents,
        categories: normalizeCategories(parsedCategories),
        periods: normalizePeriods(parsedPeriods, parsedEvents, normalizedSettings.startYear)
    };
}

function databaseToWorkbook() {
    const savedAt = new Date().toISOString();
    database.settings = { ...settings, lastActiveStartYear: activeStartYear, savedAt };
    database.categories = categories;
    database.periods = periods;
    const workbook = XLSX.utils.book_new();
    const settingRows = [["key", "value"], ...Object.entries(database.settings).map(([key, value]) => [key, value])];
    const eventRows = database.events.map((event) => ({
        id: event.id,
        academicYear: event.academicYear || inferAcademicYear(event.start),
        title: event.title,
        start: event.start,
        end: event.end,
        type: event.type,
        isNonEffective: event.isNonEffective ? "TRUE" : "FALSE",
        notes: event.notes || ""
    }));
    const categoryRows = Object.entries(categories).map(([key, category]) => ({
        key,
        label: category.label,
        color: category.color,
        textColor: category.textColor,
        defaultNonEffective: category.defaultNonEffective ? "TRUE" : "FALSE"
    }));
    const periodRows = periods.map((period) => ({
        startYear: period.startYear,
        label: period.label || `${period.startYear}/${period.startYear + 1}`,
        isDefault: period.isDefault ? "TRUE" : "FALSE",
        createdAt: period.createdAt || "",
        updatedAt: period.updatedAt || ""
    }));
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(settingRows), "Settings");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(eventRows), "Events");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(categoryRows), "Categories");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(periodRows), "Periods");
    return workbook;
}

async function saveDatabase({ manual = false } = {}) {
    const workbook = databaseToWorkbook();
    const data = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    try {
        if (dbFileHandle && await verifyPermission(dbFileHandle, true, manual)) {
            const writable = await dbFileHandle.createWritable();
            await writable.write(data);
            await writable.close();
            setDatabaseDirty(false);
            if (manual) toast("Database Excel tersimpan");
            return;
        }
    } catch (error) {
        if (error && error.name === "AbortError") return;
        console.warn(error);
    }
    if (manual && window.showSaveFilePicker) {
        try {
            dbFileHandle = await window.showSaveFilePicker({
                suggestedName: dbFileName || "kalender_database.xlsx",
                types: [{ description: "Excel Database", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }]
            });
            await storeDatabaseHandle(dbFileHandle);
            const writable = await dbFileHandle.createWritable();
            await writable.write(data);
            await writable.close();
            dbFileName = dbFileHandle.name || dbFileName;
            setDatabaseDirty(false);
            toast("Database Excel tersimpan");
            return;
        } catch (error) {
            if (error && error.name === "AbortError") return;
            console.warn(error);
        }
    }
    if (manual) {
        downloadDatabase();
        toast("Browser belum memberi izin tulis langsung. Database diunduh sebagai cadangan.", "info");
        setDatabaseDirty(false);
    }
}

async function attemptDatabaseRegeneration() {
    if (!databaseMissingOnBoot) return;
    if (!window.XLSX) return;
    if (dbFileHandle) {
        await saveDatabase({ manual: false });
        toast("Database Excel dibuat ulang dari data tersimpan.");
        return;
    }
    const guardKey = "siwaka.databaseAutoDownload.v1";
    try {
        if (sessionStorage.getItem(guardKey)) return;
        sessionStorage.setItem(guardKey, "1");
    } catch (error) {
        // Continue; the download is still the safest fallback when direct write is unavailable.
    }
    downloadDatabase();
    setDatabaseDirty(false);
    const sourceText = restoredFromLocalSnapshot ? "snapshot localStorage" : "database kosong";
    toast(`Database Excel tidak ditemukan. File baru dibuat dari ${sourceText} dan diunduh sebagai cadangan.`, "info");
}

function downloadDatabase() {
    const workbook = databaseToWorkbook();
    const data = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    downloadBlob(data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", dbFileName || "kalender_database.xlsx");
}

function openHandleDatabase() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(HANDLE_DB_NAME, 1);
        request.onupgradeneeded = () => {
            request.result.createObjectStore(HANDLE_STORE_NAME);
        };
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

async function getStoredValue(key) {
    const db = await openHandleDatabase();
    return new Promise((resolve, reject) => {
        const transaction = db.transaction(HANDLE_STORE_NAME, "readonly");
        const request = transaction.objectStore(HANDLE_STORE_NAME).get(key);
        request.onsuccess = () => resolve(request.result || null);
        request.onerror = () => reject(request.error);
        transaction.oncomplete = () => db.close();
    });
}

async function storeDatabaseHandle(handle) {
    if (!window.indexedDB) return;
    const db = await openHandleDatabase();
    await new Promise((resolve, reject) => {
        const transaction = db.transaction(HANDLE_STORE_NAME, "readwrite");
        const store = transaction.objectStore(HANDLE_STORE_NAME);
        const request = handle ? store.put(handle, DATABASE_HANDLE_KEY) : store.delete(DATABASE_HANDLE_KEY);
        request.onerror = () => reject(request.error);
        transaction.oncomplete = () => resolve();
        transaction.onerror = () => reject(transaction.error);
    });
    db.close();
}

async function verifyPermission(handle, write = false, request = true) {
    if (!handle.queryPermission) return true;
    const options = write ? { mode: "readwrite" } : { mode: "read" };
    if (await handle.queryPermission(options) === "granted") return true;
    if (!request || !handle.requestPermission) return false;
    return await handle.requestPermission(options) === "granted";
}

function markDataChanged() {
    database.settings = { ...settings };
    database.categories = categories;
    database.periods = periods;
    saveLocalDataSnapshot();
    saveUiState();
    setDatabaseDirty(true);
    if (!booting && dbFileHandle) queueAutoSave();
}

function queueAutoSave() {
    clearTimeout(autoSaveTimer);
    autoSaveTimer = setTimeout(() => saveDatabase({ manual: false }), 400);
}

function setDatabaseDirty(value) {
    dbDirty = value;
    updateDatabaseStatus(value ? "Perlu disimpan" : "Tersimpan", value ? "dirty" : "ready");
}

function updateDatabaseStatus(text, state = "") {
    $("#database-status").textContent = `Database: ${text}`;
    $("#db-file-name").textContent = dbFileName;
    const pill = $("#db-state-pill");
    pill.textContent = text;
    pill.className = `state-pill ${state}`;
}

async function exportPDF(elementId, filename, orientation = "landscape") {
    const source = document.getElementById(elementId);
    if (!source) return;
    const jsPDFCtor = window.jspdf && window.jspdf.jsPDF;
    if (!window.html2canvas || !jsPDFCtor) {
        toast("Library PDF belum tersedia.", "error");
        return;
    }
    const savedOptions = getSavedExportOptions("pdf", elementId);
    const options = await promptPdfOptions({ ...savedOptions, orientation: savedOptions.orientation || orientation }, source);
    if (!options) return;
    saveExportOptions("pdf", options, elementId);
    const pages = cloneReportPages(source, { ...options, exportKind: "pdf" });
    const wrapper = document.createElement("div");
    wrapper.className = "export-clone";
    wrapper.style.width = getPdfCloneWidth(options);
    pages.forEach((page) => wrapper.appendChild(page));
    document.body.appendChild(wrapper);
    try {
        Swal.fire({
            title: `Menyiapkan PDF ${options.paper.toUpperCase()}`,
            text: "Dokumen dibuat rapi sesuai format cetak.",
            allowOutsideClick: false,
            didOpen: () => Swal.showLoading()
        });
        const pdf = new jsPDFCtor({ unit: "mm", format: options.paper, orientation: options.orientation });
        for (let i = 0; i < pages.length; i++) {
            if (i > 0) pdf.addPage(options.paper, options.orientation);
            await addPageToPdf(pdf, pages[i], options);
        }
        pdf.save(`${sanitizeFileName(filename)}_${sanitizeFileName(settings.schoolName)}_${activeStartYear}.pdf`);
        Swal.close();
    } catch (error) {
        Swal.close();
        console.error(error);
        toast("PDF gagal dibuat.", "error");
    } finally {
        wrapper.remove();
    }
}

async function promptPdfOptions(defaults = {}, source = null) {
    const values = { ...DEFAULT_PDF_OPTIONS, ...defaults };
    if (!window.Swal) return values;
    const result = await Swal.fire({
        title: "Pengaturan Ekspor PDF",
        width: "900px",
        html: `
            <div class="export-dialog-grid pdf-export-form">
                <div class="export-dialog-fields pdf-export-fields">
                    <div class="export-dialog-intro">
                        <strong>Setelan Cetak PDF</strong>
                        <span>Atur kertas, orientasi, margin, dan skala sambil melihat pratinjau.</span>
                    </div>
                    <label>Ukuran Kertas
                        <select id="pdf-paper">
                            <option value="a3" ${values.paper === "a3" ? "selected" : ""}>A3</option>
                            <option value="a4" ${values.paper === "a4" ? "selected" : ""}>A4</option>
                        </select>
                    </label>
                    <label>Orientasi
                        <select id="pdf-orientation">
                            <option value="landscape" ${values.orientation === "landscape" ? "selected" : ""}>Horizontal / Landscape</option>
                            <option value="portrait" ${values.orientation === "portrait" ? "selected" : ""}>Vertikal / Portrait</option>
                        </select>
                    </label>
                    <label>Margin (mm)
                        <input id="pdf-margin" type="number" min="0" max="30" step="1" value="${values.margin}">
                    </label>
                    <label>Skala Isi (%)
                        <input id="pdf-fit-percent" type="number" min="60" max="120" step="1" value="${values.fitPercent}">
                    </label>
                    <label>Kualitas Render
                        <select id="pdf-capture-scale">
                            <option value="1.5" ${values.captureScale === 1.5 ? "selected" : ""}>Standar</option>
                            <option value="2" ${values.captureScale === 2 ? "selected" : ""}>Tinggi</option>
                            <option value="3" ${values.captureScale === 3 ? "selected" : ""}>Sangat Tinggi</option>
                        </select>
                    </label>
                </div>
                <div class="pdf-preview-panel">
                    <canvas id="pdf-preview-canvas" width="520" height="368" aria-label="Pratinjau hasil ekspor PDF"></canvas>
                    <p id="pdf-preview-note">Memuat pratinjau...</p>
                </div>
            </div>
        `,
        showCancelButton: true,
        confirmButtonText: "Ekspor PDF",
        cancelButtonText: "Batal",
        focusConfirm: false,
        didOpen: () => {
            const controls = ["pdf-paper", "pdf-orientation", "pdf-margin", "pdf-fit-percent"].map((id) => document.getElementById(id)).filter(Boolean);
            let previewTimer = null;
            const refresh = () => {
                clearTimeout(previewTimer);
                previewTimer = setTimeout(() => renderPdfPreview(source, readPdfOptionsFromForm()), 160);
            };
            controls.forEach((control) => {
                control.addEventListener("input", refresh);
                control.addEventListener("change", refresh);
            });
            refresh();
        },
        preConfirm: () => ({
            ...readPdfOptionsFromForm(),
            captureScale: clampNumber(document.getElementById("pdf-capture-scale").value, 1, 3, DEFAULT_PDF_OPTIONS.captureScale)
        })
    });
    return result.isConfirmed ? result.value : null;
}

function readPdfOptionsFromForm() {
    return {
        paper: document.getElementById("pdf-paper")?.value || DEFAULT_PDF_OPTIONS.paper,
        orientation: document.getElementById("pdf-orientation")?.value || DEFAULT_PDF_OPTIONS.orientation,
        margin: clampNumber(document.getElementById("pdf-margin")?.value, 0, 30, DEFAULT_PDF_OPTIONS.margin),
        fitPercent: clampNumber(document.getElementById("pdf-fit-percent")?.value, 60, 120, DEFAULT_PDF_OPTIONS.fitPercent),
        captureScale: clampNumber(document.getElementById("pdf-capture-scale")?.value, 1, 3, DEFAULT_PDF_OPTIONS.captureScale)
    };
}

async function renderPdfPreview(source, options) {
    const canvas = document.getElementById("pdf-preview-canvas");
    if (!canvas || !source || !window.html2canvas) return;
    const requestId = ++pdfPreviewRequestId;
    const note = document.getElementById("pdf-preview-note");
    if (note) note.textContent = "Memuat pratinjau...";
    const pages = cloneReportPages(source, { ...options, exportKind: "pdf" });
    const page = pages[0];
    const wrapper = document.createElement("div");
    wrapper.className = "export-clone pdf-preview-clone";
    wrapper.style.width = getPdfCloneWidth(options);
    wrapper.appendChild(page);
    document.body.appendChild(wrapper);
    try {
        await new Promise((resolve) => requestAnimationFrame(resolve));
        if (document.fonts?.ready) await document.fonts.ready;
        const rendered = await html2canvas(page, {
            scale: 0.35,
            backgroundColor: "#ffffff",
            useCORS: true,
            logging: false,
            windowWidth: Math.ceil(page.scrollWidth),
            windowHeight: Math.ceil(page.scrollHeight),
            letterRendering: true
        });
        if (requestId !== pdfPreviewRequestId) return;
        drawPdfPreview(canvas, rendered, options);
        if (note) note.textContent = `${options.paper.toUpperCase()} ${options.orientation === "landscape" ? "Horizontal" : "Vertikal"} - margin ${options.margin} mm - skala ${options.fitPercent}%`;
    } catch (error) {
        if (note) note.textContent = "Pratinjau belum dapat dibuat.";
    } finally {
        wrapper.remove();
    }
}

function drawPdfPreview(canvas, rendered, options) {
    const context = canvas.getContext("2d");
    const pageSize = getPdfPaperSize(options);
    const canvasWidth = 520;
    const canvasHeight = Math.round(canvasWidth * (pageSize.height / pageSize.width));
    canvas.width = canvasWidth;
    canvas.height = canvasHeight;
    context.fillStyle = "#e5e7eb";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.fillStyle = "#ffffff";
    context.fillRect(8, 8, canvas.width - 16, canvas.height - 16);
    context.strokeStyle = "#94a3b8";
    context.lineWidth = 1;
    context.strokeRect(8, 8, canvas.width - 16, canvas.height - 16);
    const paperW = canvas.width - 16;
    const paperH = canvas.height - 16;
    const unit = paperW / pageSize.width;
    const margin = options.margin * unit;
    const availableW = paperW - margin * 2;
    const availableH = paperH - margin * 2;
    const ratio = Math.min(availableW / rendered.width, availableH / rendered.height) * (options.fitPercent / 100);
    const drawW = rendered.width * ratio;
    const drawH = rendered.height * ratio;
    const x = 8 + (paperW - drawW) / 2;
    const y = 8 + (paperH - drawH) / 2;
    context.drawImage(rendered, x, y, drawW, drawH);
}

function getPdfPaperSize(options) {
    const base = options.paper === "a4" ? { width: 210, height: 297 } : { width: 297, height: 420 };
    return options.orientation === "landscape"
        ? { width: base.height, height: base.width }
        : base;
}

function getPdfCloneWidth(options) {
    if (options.paper === "a4") return options.orientation === "landscape" ? "1400px" : "980px";
    return options.orientation === "landscape" ? "1900px" : "1340px";
}

async function addPageToPdf(pdf, page, options) {
    await new Promise((resolve) => requestAnimationFrame(resolve));
    if (document.fonts?.ready) await document.fonts.ready;
    const canvas = await html2canvas(page, {
        scale: options.captureScale,
        backgroundColor: "#ffffff",
        useCORS: true,
        logging: false,
        windowWidth: Math.ceil(page.scrollWidth),
        windowHeight: Math.ceil(page.scrollHeight),
        letterRendering: true,
        removeContainer: true
    });
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const margin = options.margin;
    const ratio = Math.min((pageWidth - margin * 2) / canvas.width, (pageHeight - margin * 2) / canvas.height) * (options.fitPercent / 100);
    const renderWidth = canvas.width * ratio;
    const renderHeight = canvas.height * ratio;
    const x = (pageWidth - renderWidth) / 2;
    const y = (pageHeight - renderHeight) / 2;
    pdf.addImage(canvas.toDataURL("image/png"), "PNG", x, y, renderWidth, renderHeight);
}

async function exportExcel(elementId, filename) {
    const source = document.getElementById(elementId);
    if (!source) return;
    if (!window.ExcelJS) {
        toast("Library Excel belum tersedia.", "error");
        return;
    }
    if (elementId === "export-dashboard") {
        await exportDashboardExcel(filename);
        return;
    }
    if (elementId === "export-efektif") {
        await exportEffectiveExcel(filename);
        return;
    }
    if (elementId === "export-rekap") {
        await exportRekapExcel(filename);
        return;
    }
    const pages = cloneReportPages(source);
    const wrapper = document.createElement("div");
    wrapper.className = "export-clone";
    pages.forEach((page) => wrapper.appendChild(page));
    document.body.appendChild(wrapper);
    try {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "SiWaka";
        pages.forEach((page, index) => {
            const sheetName = sanitizeSheetName(page.dataset.sheetName || page.querySelector(".format-note")?.textContent || `Laporan ${index + 1}`);
            const sheet = workbook.addWorksheet(uniqueExcelSheetName(workbook, sheetName));
            sheet.pageSetup = { paperSize: 8, orientation: page.classList.contains("effective-sheet") || elementId === "export-rekap" ? "portrait" : "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1, margins: { left: 0.25, right: 0.25, top: 0.35, bottom: 0.35, header: 0.1, footer: 0.1 } };
            addPageToExcelSheet(sheet, page);
        });
        const buffer = await workbook.xlsx.writeBuffer();
        downloadBlob(buffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", `${sanitizeFileName(filename)}_${sanitizeFileName(settings.schoolName)}_${activeStartYear}.xlsx`);
    } catch (error) {
        console.error(error);
        toast("Excel gagal dibuat.", "error");
    } finally {
        wrapper.remove();
    }
}

async function exportDashboardExcel(filename) {
    try {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "SiWaka";
        const sheet = workbook.addWorksheet("Dashboard");
        sheet.views = [{ showGridLines: false }];
        sheet.pageSetup = { paperSize: 8, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1, horizontalCentered: true, verticalCentered: false, margins: excelMargins(0.18) };
        sheet.columns = Array.from({ length: 31 }, (_, index) => ({ width: [7, 15, 23].includes(index) ? 1.4 : 4 }));
        mergeSet(sheet, 1, 1, 1, 31, `KALENDER PENDIDIKAN ${settings.schoolName}`, titleExcelStyle(16));
        mergeSet(sheet, 2, 1, 2, 31, getDashboardSubtitle(), titleExcelStyle(14));
        const months = dashboardMode === "month"
            ? [{ m: currentDate.getMonth(), y: currentDate.getFullYear() }]
            : getAcademicMonths(dashboardMode === "year" ? "year" : dashboardMode);
        months.forEach((item, index) => {
            const blockRow = 4 + Math.floor(index / 4) * 9;
            const blockCol = 1 + (index % 4) * 8;
            addDashboardMonthToExcel(sheet, item.y, item.m, blockRow, blockCol);
        });
        const summaryRow = 4 + Math.ceil(months.length / 4) * 9 + 1;
        addAcademicSummaryToExcel(sheet, summaryRow, 2, 5, 3);
        addExcelSignature(sheet, summaryRow, 22, 31, "kepala");
        sheet.pageSetup.printArea = `A1:AE${summaryRow + 6}`;
        const buffer = await workbook.xlsx.writeBuffer();
        downloadBlob(buffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", `${sanitizeFileName(filename)}_${sanitizeFileName(settings.schoolName)}_${activeStartYear}.xlsx`);
    } catch (error) {
        console.error(error);
        toast("Excel Dashboard gagal dibuat.", "error");
    }
}

function addDashboardMonthToExcel(sheet, year, month, startRow, startCol) {
    const endCol = startCol + 6;
    mergeSet(sheet, startRow, startCol, startRow, startCol, month + 1, monthHeaderExcelStyle());
    mergeSet(sheet, startRow, startCol + 1, startRow, startCol + 5, monthNames[month], monthHeaderExcelStyle(true));
    mergeSet(sheet, startRow, endCol, startRow, endCol, year, monthHeaderExcelStyle());
    dayNames.forEach((day, index) => {
        const cell = sheet.getCell(startRow + 1, startCol + index);
        cell.value = day;
        applyCellStyle(cell, {
            fill: "FF93C5FD",
            font: { bold: true, color: index === 0 || (Number(settings.workDays) === 5 && index === 6) ? "FFFF0000" : "FF111111" },
            border: true,
            align: "center"
        });
    });
    const first = new Date(year, month, 1);
    const start = new Date(year, month, 1 - first.getDay());
    for (let week = 0; week < 6; week++) {
        for (let day = 0; day < 7; day++) {
            const date = new Date(start);
            date.setDate(start.getDate() + week * 7 + day);
            const row = startRow + 2 + week;
            const col = startCol + day;
            const cell = sheet.getCell(row, col);
            const other = date.getMonth() !== month;
            const events = other ? [] : getDayEvents(formatDate(date));
            const category = events.length ? (categories[events[0].type] || categories.lainnya) : null;
            cell.value = other ? "" : date.getDate();
            applyCellStyle(cell, {
                fill: category ? argbFromHex(category.color) : "FFFFFFFF",
                font: { bold: true, color: category ? argbTextColor(category.textColor) : !isWorkingDay(date) ? "FFFF0000" : "FF111111" },
                border: true,
                align: "center"
            });
            sheet.getRow(row).height = 22;
        }
    }
}

async function exportEffectiveExcel(filename) {
    try {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "SiWaka";
        buildEffectiveExcelSheet(workbook, "Ganjil", hitungSemester(getAcademicMonths("semester1")), "FF93C5FD", `${activeStartYear}-07-13`);
        buildEffectiveExcelSheet(workbook, "Genap", hitungSemester(getAcademicMonths("semester2")), "FFFEF08A", `${activeStartYear + 1}-01-02`);
        const buffer = await workbook.xlsx.writeBuffer();
        downloadBlob(buffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", `${sanitizeFileName(filename)}_${sanitizeFileName(settings.schoolName)}_${activeStartYear}.xlsx`);
    } catch (error) {
        console.error(error);
        toast("Excel Waktu Efektif gagal dibuat.", "error");
    }
}

function buildEffectiveExcelSheet(workbook, semesterName, result, accent, ttdDate) {
    return addEffectiveSemesterToExcel(workbook, semesterName, result, accent, ttdDate);
}

function addEffectiveSemesterToExcel(workbook, semesterName, result, accent, ttdDate) {
    const sheet = workbook.addWorksheet(`Semester ${semesterName}`);
    sheet.views = [{ showGridLines: false }];
    sheet.pageSetup = {
        paperSize: 9,
        orientation: "portrait",
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 1,
        horizontalCentered: true,
        verticalCentered: false,
        margins: { left: 0.22, right: 0.22, top: 0.28, bottom: 0.22, header: 0.08, footer: 0.08 }
    };
    sheet.pageSetup.printArea = "A1:H31";
    sheet.columns = [
        { width: 5.5 }, { width: 16 }, { width: 9.5 }, { width: 9.5 },
        { width: 11 }, { width: 12.5 }, { width: 10.5 }, { width: 9.5 }
    ];
    sheet.properties.defaultRowHeight = 17;
    mergeSet(sheet, 1, 1, 1, 8, "ANALISIS PEKAN EFEKTIF", titleExcelStyle(15));
    sheet.getRow(1).height = 24;
    sheet.getRow(2).height = 17;

    const identityRows = [
        ["Nama Sekolah", `: ${settings.schoolName}`, "Kelas", `: ${settings.className}`],
        ["Mata pelajaran", `: ${settings.subjectName}`, "Tahun Ajaran", `: ${activeStartYear}-${activeStartYear + 1}`],
        ["Nama Guru", `: ${formatNameForSignature(settings.teacherName)}`, "Semester", `: ${semesterName}`]
    ];
    identityRows.forEach((items, idx) => {
        const row = 3 + idx;
        sheet.getCell(row, 1).value = items[0];
        mergeSet(sheet, row, 2, row, 3, items[1], { align: "left", font: { color: "FF111111" } });
        sheet.getCell(row, 6).value = items[2];
        mergeSet(sheet, row, 7, row, 8, items[3], { align: "left", font: { color: "FF111111" } });
        for (let col = 1; col <= 8; col++) {
            applyCellStyle(sheet.getCell(row, col), { align: "left", font: { color: "FF111111" }, wrap: false });
        }
        sheet.getRow(row).height = 18;
    });

    mergeSet(sheet, 7, 1, 7, 8, "A. Pekan Efektif", { font: { bold: true, color: "FF111111" }, align: "left" });
    sheet.getRow(6).height = 15;
    sheet.getRow(7).height = 21;
    addEffectiveMainTableToExcel(sheet, result, 8, accent);

    const summaryTitleRow = 18;
    mergeSet(sheet, summaryTitleRow, 1, summaryTitleRow, 8, "B. Jumlah Jam Tatap Muka", { font: { bold: true, color: "FF111111" }, align: "left" });
    sheet.getRow(summaryTitleRow).height = 20;
    const total = result.total;
    [
        ["a. Jumlah pekan efektif KBM", ":", total.pekanEfektif, "Pekan"],
        ["b. Pekan tidak efektif KBM", ":", total.pekanTidakEfektif, "Pekan"],
        ["c. Jumlah hari efektif KBM", ":", total.hariEfektif, "Hari"],
        ["d. Alokasi jam per minggu", ":", settings.jamPerMinggu, "Jam"],
        ["e. Jumlah jam Efektif", ":", total.jamEfektif, "Jam"]
    ].forEach((rowValues, index) => {
        const row = summaryTitleRow + 1 + index;
        mergeSet(sheet, row, 2, row, 4, rowValues[0], { align: "left", font: { color: "FF111111" } });
        sheet.getCell(row, 5).value = rowValues[1];
        sheet.getCell(row, 6).value = rowValues[2];
        sheet.getCell(row, 7).value = rowValues[3];
        for (let col = 1; col <= 8; col++) {
            applyCellStyle(sheet.getCell(row, col), { align: [5, 6].includes(col) ? "center" : "left", font: { color: "FF111111" }, wrap: false });
        }
        sheet.getRow(row).height = 19;
    });

    sheet.getRow(24).height = 15;
    addExcelSignature(sheet, 25, 2, 3, "kepala", "Mengetahui", "Kepala Sekolah");
    addExcelSignature(sheet, 25, 6, 8, "guru", `${settings.ttdPlace}, ${formatDateIndo(ttdDate)}`, "Guru Mata Pelajaran");
    for (let row = 25; row <= 31; row++) sheet.getRow(row).height = row === 27 || row === 28 ? 28 : 19;
}

function addEffectiveMainTableToExcel(sheet, result, startRow, accent) {
    const headers = [
        ["No", "Bulan", "Tersedia", "", "Waktu Tidak Efektif", "", "Waktu Efektif", ""],
        ["", "", "Pekan", "Hari", "Libur Akhir Pekan", "Libur Hari Lainnya", "Hari", "Pekan"]
    ];
    setExcelRowValues(sheet, startRow, headers[0]);
    setExcelRowValues(sheet, startRow + 1, headers[1]);
    sheet.mergeCells(startRow, 1, startRow + 1, 1);
    sheet.mergeCells(startRow, 2, startRow + 1, 2);
    sheet.mergeCells(startRow, 3, startRow, 4);
    sheet.mergeCells(startRow, 5, startRow, 6);
    sheet.mergeCells(startRow, 7, startRow, 8);
    for (let row = startRow; row <= startRow + 1; row++) {
        for (let col = 1; col <= 8; col++) applyCellStyle(sheet.getCell(row, col), { fill: accent, font: { bold: true, color: "FF111111" }, border: true, align: "center" });
    }
    sheet.getRow(startRow).height = 24;
    sheet.getRow(startRow + 1).height = 36;
    result.data.forEach((item, index) => {
        const row = startRow + 2 + index;
        setExcelRowValues(sheet, row, [item.no, item.bulan, item.pekanTersedia, item.hariDalamBulan, item.hariNonKerja, item.hariTidakEfektif, item.hariEfektif, item.pekanEfektif]);
        for (let col = 1; col <= 8; col++) applyCellStyle(sheet.getCell(row, col), { border: true, align: col === 2 ? "left" : "center" });
        sheet.getRow(row).height = item.bulan.length > 12 ? 28 : 22;
    });
    const totalRow = startRow + 2 + result.data.length;
    const t = result.total;
    setExcelRowValues(sheet, totalRow, ["", "", t.pekanTersedia, t.hariDalamBulan, t.hariNonKerja, t.hariTidakEfektif, t.hariEfektif, t.pekanEfektif]);
    for (let col = 1; col <= 8; col++) applyCellStyle(sheet.getCell(totalRow, col), { fill: accent, font: { bold: true, color: "FF111111" }, border: true, align: "center" });
    sheet.getRow(totalRow).height = 22;
}

function setExcelRowValues(sheet, rowNumber, values) {
    values.forEach((value, index) => {
        sheet.getCell(rowNumber, index + 1).value = value;
    });
}

async function exportRekapExcel(filename) {
    try {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "SiWaka";
        buildRekapExcelSheet(workbook);
        const buffer = await workbook.xlsx.writeBuffer();
        downloadBlob(buffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", `${sanitizeFileName(filename)}_${sanitizeFileName(settings.schoolName)}_${activeStartYear}.xlsx`);
    } catch (error) {
        console.error(error);
        toast("Excel Rekapitulasi gagal dibuat.", "error");
    }
}

function buildRekapExcelSheet(workbook) {
    const sheet = workbook.addWorksheet("Rekapitulasi");
    sheet.views = [{ showGridLines: false }];
    sheet.pageSetup = { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 1, margins: excelMargins(0.25) };
    sheet.columns = [{ width: 4 }, { width: 34 }, { width: 15 }, { width: 15 }, { width: 8 }, { width: 25 }, { width: 20 }];
    mergeSet(sheet, 1, 1, 1, 7, "REKAPITULASI KEGIATAN DAN HARI LIBUR", titleExcelStyle(14));
    mergeSet(sheet, 2, 1, 2, 7, `TAHUN PELAJARAN ${activeStartYear}-${activeStartYear + 1}`, titleExcelStyle(11));
    const filtered = filterRekapEvents();
    let row = 4;
    row = addRekapSectionToExcel(sheet, row, "A. KEGIATAN TETAP KBM", filtered.filter((event) => !event.isNonEffective), "Tetap KBM");
    row += 2;
    row = addRekapSectionToExcel(sheet, row, "B. HARI TIDAK EFEKTIF KBM", filtered.filter((event) => event.isNonEffective), "Hari Tidak Efektif KBM");
    addExcelSignature(sheet, row + 2, 5, 7, "kepala", `${settings.ttdPlace}, ${formatDateIndo(settings.ttdDate)}`, "Kepala Sekolah");
    return sheet;
}

function addRekapSectionToExcel(sheet, startRow, title, events, status) {
    mergeSet(sheet, startRow, 1, startRow, 7, title, { font: { bold: true, color: "FF111111" }, align: "left" });
    const headerRow = startRow + 1;
    ["No", "Nama Kegiatan", "Tanggal Mulai", "Tanggal Berakhir", "Durasi", "Kategori", "Status KBM"].forEach((header, index) => {
        const cell = sheet.getCell(headerRow, index + 1);
        cell.value = header;
        applyCellStyle(cell, { fill: "FFE8EDF5", font: { bold: true, color: "FF111111" }, border: true, align: "center" });
    });
    if (!events.length) {
        mergeSet(sheet, headerRow + 1, 1, headerRow + 1, 7, "Tidak ada data.", { border: true, align: "center", font: { color: "FF667085" } });
        return headerRow + 2;
    }
    events.forEach((event, index) => {
        const row = headerRow + 1 + index;
        const category = categories[event.type] || categories.lainnya;
        const values = [index + 1, event.title, formatDateIndo(event.start), formatDateIndo(event.end), `${calculateDays(event.start, event.end)} Hari`, category.label, status];
        values.forEach((value, colIndex) => {
            const cell = sheet.getCell(row, colIndex + 1);
            cell.value = value;
            const isCategory = colIndex === 5;
            const isStatus = colIndex === 6;
            applyCellStyle(cell, {
                fill: isCategory ? argbFromHex(category.color) : "FFFFFFFF",
                font: {
                    bold: isCategory || isStatus,
                    color: isCategory ? argbTextColor(category.textColor) : isStatus ? (event.isNonEffective ? "FFFF0000" : "FF008000") : "FF111111"
                },
                border: true,
                align: [1, 2, 3].includes(colIndex) ? "left" : "center"
            });
        });
    });
    return headerRow + events.length + 1;
}

function addAcademicSummaryToExcel(sheet, startRow, startCol, labelSpan = 5, valueSpan = 3) {
    const semester1 = hitungSemester(getAcademicMonths("semester1"));
    const semester2 = hitungSemester(getAcademicMonths("semester2"));
    const rows = [
        ["", "Semester I", "Semester II"],
        ["Hari Efektif KBM", `${semester1.total.hariEfektif} Hari`, `${semester2.total.hariEfektif} Hari`],
        ["Pekan Efektif KBM", `${semester1.total.pekanEfektif} Pekan`, `${semester2.total.pekanEfektif} Pekan`],
        ["Pekan Tidak Efektif KBM", `${semester1.total.pekanTidakEfektif} Pekan`, `${semester2.total.pekanTidakEfektif} Pekan`]
    ];
    rows.forEach((items, index) => {
        const row = startRow + index;
        items.forEach((value, offset) => {
            const col = offset === 0
                ? startCol
                : startCol + labelSpan + (offset - 1) * valueSpan;
            const span = offset === 0 ? labelSpan : valueSpan;
            sheet.mergeCells(row, col, row, col + span - 1);
            const cell = sheet.getCell(row, col);
            cell.value = value;
            applyCellStyle(cell, { fill: index === 0 ? "FFE8EDF5" : "FFFFFFFF", font: { bold: index === 0, color: "FF111111" }, border: true, align: offset === 0 && index > 0 ? "left" : "center" });
        });
    });
}

function excelMargins(value) {
    return { left: value, right: value, top: value, bottom: value, header: 0.1, footer: 0.1 };
}

function mergeSet(sheet, startRow, startCol, endRow, endCol, value, style = {}) {
    if (startRow !== endRow || startCol !== endCol) sheet.mergeCells(startRow, startCol, endRow, endCol);
    const cell = sheet.getCell(startRow, startCol);
    cell.value = value;
    applyCellStyle(cell, style);
    return cell;
}

function titleExcelStyle(size = 14) {
    return { font: { bold: true, size, color: "FF000000" }, align: "center" };
}

function monthHeaderExcelStyle(bold = false) {
    return { fill: "FF93C5FD", font: { bold, color: "FF111111" }, border: true, align: "center" };
}

function applyCellStyle(cell, options = {}) {
    if (options.fill) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: options.fill } };
    const font = options.font || {};
    cell.font = {
        bold: Boolean(font.bold),
        underline: Boolean(font.underline),
        size: font.size || 10,
        color: { argb: font.color || "FF111111" }
    };
    cell.alignment = {
        horizontal: options.align || "center",
        vertical: "middle",
        wrapText: options.wrap !== false
    };
    if (options.border) {
        cell.border = {
            top: { style: "thin", color: { argb: "FF111111" } },
            left: { style: "thin", color: { argb: "FF111111" } },
            bottom: { style: "thin", color: { argb: "FF111111" } },
            right: { style: "thin", color: { argb: "FF111111" } }
        };
    }
}

function argbFromHex(value) {
    return `FF${normalizeColor(value).replace("#", "").toUpperCase()}`;
}

function argbTextColor(value) {
    return String(value).toLowerCase() === "black" ? "FF111111" : "FFFFFFFF";
}

function addExcelSignature(sheet, startRow, startCol, endCol, type = "kepala", line1 = "", line2 = "") {
    const name = type === "guru" ? settings.teacherName : settings.headmasterName;
    const nip = type === "guru" ? settings.teacherNip : settings.headmasterNip;
    const lines = [
        line1 || `${settings.ttdPlace}, ${formatDateIndo(settings.ttdDate)}`,
        line2 || (type === "guru" ? "Guru Mata Pelajaran" : `Kepala ${settings.schoolName}`),
        "",
        "",
        formatNameForSignature(name),
        `NIP. ${nip}`
    ];
    lines.forEach((text, index) => {
        const row = startRow + index;
        sheet.mergeCells(row, startCol, row, endCol);
        const cell = sheet.getCell(row, startCol);
        cell.value = text;
        applyCellStyle(cell, {
            font: { bold: index === 4, underline: index === 4, color: "FF111111" },
            align: "center"
        });
        if (index === 2 || index === 3) sheet.getRow(row).height = 28;
    });
}


function addPageToExcelSheet(sheet, page) {
    let rowIndex = 1;
    const maxCols = Math.max(...$$("table", page).map((table) => getTableMaxColumns(table)), 8);
    sheet.columns = Array.from({ length: maxCols }, () => ({ width: page.classList.contains("akademik-sheet") ? 4.8 : 14 }));
    if (page.classList.contains("akademik-sheet") && maxCols >= 2) sheet.getColumn(2).width = 15;
    sheet.properties.defaultRowHeight = page.classList.contains("akademik-sheet") ? 18 : 22;
    sheet.views = [{ showGridLines: false }];
    $$(":scope > *", page).forEach((child) => {
        rowIndex = addElementToExcel(sheet, child, rowIndex, maxCols);
    });
}

function addElementToExcel(sheet, element, rowIndex, maxCols) {
    if (element.classList && element.classList.contains("no-export")) return rowIndex;
    const tag = element.tagName ? element.tagName.toLowerCase() : "";
    if (element.matches?.(".signature-row")) return addSignatureRowToExcel(sheet, element, rowIndex, maxCols) + 2;
    if (element.matches?.(".signature-block")) return addSingleSignatureBlockToExcel(sheet, element, rowIndex, maxCols) + 2;
    if (tag === "table") return addHtmlTableToExcel(sheet, element, rowIndex) + 2;
    const directTable = element.matches && element.matches(".table-scroll") ? element.querySelector("table") : null;
    if (directTable) return addHtmlTableToExcel(sheet, directTable, rowIndex) + 2;
    if (["h1", "h2", "h3", "p"].includes(tag) || element.classList?.contains("effective-title")) {
        const text = cleanText(element.innerText || element.textContent);
        if (text) {
            const row = sheet.getRow(rowIndex);
            row.getCell(1).value = text;
            row.getCell(1).font = { bold: tag !== "p", size: tag === "h2" || element.classList?.contains("effective-title") ? 16 : 12 };
            row.getCell(1).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
            sheet.mergeCells(rowIndex, 1, rowIndex, Math.max(1, maxCols));
            row.height = 24;
            rowIndex += 1;
        }
        return rowIndex;
    }
    const children = $$(":scope > *", element).filter((node) => !node.classList?.contains("no-export"));
    if (children.length) {
        children.forEach((child) => { rowIndex = addElementToExcel(sheet, child, rowIndex, maxCols); });
        return rowIndex;
    }
    return rowIndex;
}

function addSignatureRowToExcel(sheet, rowNode, startRow, maxCols) {
    const blocks = $$(".signature-block", rowNode);
    const leftEnd = Math.max(1, Math.floor(maxCols / 2));
    const ranges = [
        { start: 1, end: leftEnd, block: blocks[0] },
        { start: leftEnd + 1, end: maxCols, block: blocks[1] }
    ];
    const lines = ranges.map((range) => signatureLines(range.block));
    const rowCount = Math.max(...lines.map((items) => items.length), 1);
    for (let offset = 0; offset < rowCount; offset++) {
        const excelRow = sheet.getRow(startRow + offset);
        ranges.forEach((range, index) => {
            const line = lines[index][offset] || { text: "" };
            writeSignatureLine(sheet, excelRow, startRow + offset, range.start, range.end, line);
        });
        if (lines.some((items) => items[offset]?.space)) excelRow.height = 54;
    }
    return startRow + rowCount - 1;
}

function addSingleSignatureBlockToExcel(sheet, block, startRow, maxCols) {
    const startCol = Math.max(1, Math.floor(maxCols / 2) + 1);
    const lines = signatureLines(block);
    lines.forEach((line, index) => {
        const rowIndex = startRow + index;
        const excelRow = sheet.getRow(rowIndex);
        writeSignatureLine(sheet, excelRow, rowIndex, startCol, maxCols, line);
        if (line.space) excelRow.height = 54;
    });
    return startRow + lines.length - 1;
}

function signatureLines(block) {
    if (!block) return [{ text: "" }];
    const lines = [];
    $$(":scope > *", block).forEach((child) => {
        if (child.classList.contains("signature-space")) {
            lines.push({ text: "", space: true });
            return;
        }
        const text = cleanText(child.innerText || child.textContent);
        if (!text) return;
        lines.push({
            text,
            bold: child.classList.contains("signature-name"),
            underline: child.classList.contains("signature-name")
        });
    });
    return lines.length ? lines : [{ text: "" }];
}

function writeSignatureLine(sheet, excelRow, rowIndex, startCol, endCol, line) {
    if (startCol <= endCol && startCol !== endCol) {
        try {
            sheet.mergeCells(rowIndex, startCol, rowIndex, endCol);
        } catch (error) {
            // The cell may already be merged by a parallel signature block on a narrow sheet.
        }
    }
    const cell = excelRow.getCell(startCol);
    cell.value = line.text || "";
    cell.font = { bold: Boolean(line.bold), underline: Boolean(line.underline), size: 11, color: { argb: "FF111111" } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
}

function addHtmlTableToExcel(sheet, table, startRow) {
    const occupied = {};
    let rowIndex = startRow;
    Array.from(table.rows).forEach((htmlRow) => {
        const excelRow = sheet.getRow(rowIndex);
        let colIndex = 1;
        Array.from(htmlRow.cells).forEach((cell) => {
            while (occupied[`${rowIndex}:${colIndex}`]) colIndex++;
            const rowspan = Number(cell.getAttribute("rowspan") || 1);
            const colspan = Number(cell.getAttribute("colspan") || 1);
            const excelCell = excelRow.getCell(colIndex);
            excelCell.value = cleanText(cell.innerText || cell.textContent);
            applyExcelCellStyle(excelCell, cell);
            if (rowspan > 1 || colspan > 1) {
                sheet.mergeCells(rowIndex, colIndex, rowIndex + rowspan - 1, colIndex + colspan - 1);
                for (let r = rowIndex; r < rowIndex + rowspan; r++) {
                    for (let c = colIndex; c < colIndex + colspan; c++) occupied[`${r}:${c}`] = true;
                }
            }
            colIndex += colspan;
        });
        excelRow.height = table.classList.contains("academic-board") ? 18 : 24;
        rowIndex++;
    });
    return rowIndex - 1;
}

function applyExcelCellStyle(excelCell, htmlCell) {
    const style = getComputedStyle(htmlCell);
    const fill = getCellBackgroundArgb(htmlCell);
    const fontColor = getCellFontArgb(htmlCell) || "FF111827";
    const isAcademic = Boolean(htmlCell.closest(".academic-board"));
    excelCell.font = { bold: htmlCell.tagName.toLowerCase() === "th" || Number(style.fontWeight) >= 600, color: { argb: fontColor }, size: isAcademic ? 9 : 11 };
    excelCell.alignment = { horizontal: style.textAlign === "left" ? "left" : "center", vertical: "middle", wrapText: !isAcademic, shrinkToFit: isAcademic };
    if (!htmlCell.closest(".effective-summary") && !htmlCell.closest(".effective-identity-table")) {
        excelCell.border = {
            top: { style: "thin", color: { argb: "FF111111" } },
            left: { style: "thin", color: { argb: "FF111111" } },
            bottom: { style: "thin", color: { argb: "FF111111" } },
            right: { style: "thin", color: { argb: "FF111111" } }
        };
    }
    if (fill && fill !== "00000000" && fill !== "FFFFFFFF") {
        excelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fill } };
    }
}

function getCellBackgroundArgb(htmlCell) {
    const style = getComputedStyle(htmlCell);
    const direct = cssColorToArgb(style.backgroundColor);
    if (direct && direct !== "00000000" && direct !== "FFFFFFFF") return direct;
    const inline = extractFirstHexColor(htmlCell.getAttribute("style") || "");
    if (inline) return cssColorToArgb(inline);
    const coloredChild = htmlCell.querySelector(".category-pill, .event-mini, .swatch");
    if (coloredChild) {
        const childStyle = getComputedStyle(coloredChild);
        return cssColorToArgb(childStyle.backgroundColor) || cssColorToArgb(extractFirstHexColor(coloredChild.getAttribute("style") || ""));
    }
    return direct;
}

function getCellFontArgb(htmlCell) {
    const coloredChild = htmlCell.querySelector(".category-pill, .event-mini");
    if (coloredChild) return cssColorToArgb(getComputedStyle(coloredChild).color);
    return cssColorToArgb(getComputedStyle(htmlCell).color);
}

function extractFirstHexColor(value) {
    const match = String(value || "").match(/#[0-9a-f]{6}/i);
    return match ? match[0] : "";
}

async function exportWord(elementId, filename) {
    const source = document.getElementById(elementId);
    if (!source) return;
    if (!window.docx) {
        toast("Library Word belum tersedia.", "error");
        return;
    }
    const options = getWordOptionsForTarget(elementId);
    if (elementId === "export-dashboard") {
        await exportDashboardWord(filename, options);
        return;
    }
    const pages = cloneReportPages(source, options);
    const wrapper = document.createElement("div");
    wrapper.className = "export-clone";
    pages.forEach((page) => wrapper.appendChild(page));
    document.body.appendChild(wrapper);
    try {
        const sections = pages.map((page) => ({
            properties: {
                page: {
                    size: getDocxPageSize(options.orientation === "portrait"),
                    margin: docxMarginFromMm(options.margin)
                }
            },
            children: pageToDocxChildren(page)
        }));
        const document = new docx.Document({ sections });
        const blob = await docx.Packer.toBlob(document);
        downloadBlob(blob, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", `${sanitizeFileName(filename)}_${sanitizeFileName(settings.schoolName)}_${activeStartYear}.docx`);
    } catch (error) {
        console.error(error);
        toast("Word gagal dibuat.", "error");
    } finally {
        wrapper.remove();
    }
}

function getWordOptionsForTarget(elementId) {
    const portrait = elementId === "export-efektif" || elementId === "export-rekap";
    return {
        ...DEFAULT_WORD_OPTIONS,
        orientation: portrait ? "portrait" : "landscape",
        margin: portrait ? 10 : 5,
        fitMode: "compact"
    };
}

async function exportDashboardWord(filename, options) {
    try {
        const months = dashboardMode === "month"
            ? [{ m: currentDate.getMonth(), y: currentDate.getFullYear() }]
            : getAcademicMonths(dashboardMode === "year" ? "year" : dashboardMode);
        const children = [
            new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 80 },
                children: [new docx.TextRun({ text: `KALENDER PENDIDIKAN ${settings.schoolName}`, bold: true, size: 28 })]
            }),
            new docx.Paragraph({
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 160 },
                children: [new docx.TextRun({ text: getDashboardSubtitle(), bold: true, size: 22 })]
            }),
            dashboardMonthsToDocxGrid(months),
            summarySignatureFooterToDocx(null, { dashboardFooter: true })
        ];
        const document = new docx.Document({
            sections: [{
                properties: {
                    page: {
                        size: getDocxPageSize(options.orientation === "portrait"),
                        margin: docxMarginFromMm(options.margin)
                    }
                },
                children
            }]
        });
        const blob = await docx.Packer.toBlob(document);
        downloadBlob(blob, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", `${sanitizeFileName(filename)}_${sanitizeFileName(settings.schoolName)}_${activeStartYear}.docx`);
    } catch (error) {
        console.error(error);
        toast("Word Dashboard gagal dibuat.", "error");
    }
}

function summarySignatureFooterToDocx(block = null, options = {}) {
    const generatedBlock = htmlStringToElement(renderSignatureBlock("kepala", true));
    const signatureBlock = block || (generatedBlock.matches(".signature-block") ? generatedBlock : generatedBlock.querySelector(".signature-block"));
    const summaryWidth = options.dashboardFooter ? 44 : options.wideSummary ? 46 : 38;
    const signatureWidth = 100 - summaryWidth;
    return new docx.Table({
        width: { size: 100, type: docx.WidthType.PERCENTAGE },
        borders: noWordBorders(),
        rows: [new docx.TableRow({
            children: [
                new docx.TableCell({
                    width: { size: summaryWidth, type: docx.WidthType.PERCENTAGE },
                    borders: noWordBorders(),
                    verticalAlign: docx.VerticalAlign.TOP,
                    margins: { top: 80, bottom: 20, left: 40, right: options.dashboardFooter ? 60 : options.wideSummary ? 50 : 90 },
                    children: [compactAcademicSummaryToDocx(options)]
                }),
                new docx.TableCell({
                    width: { size: signatureWidth, type: docx.WidthType.PERCENTAGE },
                    borders: noWordBorders(),
                    verticalAlign: docx.VerticalAlign.TOP,
                    margins: { top: 70, bottom: 20, left: options.dashboardFooter ? 220 : options.wideSummary ? 70 : 110, right: 0 },
                    children: signatureBlockToDocxParagraphs(signatureBlock, docx.AlignmentType.CENTER, { size: 24 })
                })
            ]
        })]
    });
}

function compactAcademicSummaryToDocx(options = {}) {
    const headerSize = options.wideSummary ? 24 : options.dashboardFooter ? 22 : 12;
    const bodySize = options.wideSummary ? 24 : options.dashboardFooter ? 22 : 11;
    const rows = [
        new docx.TableRow({
            children: [
                wordCell("", { fill: "E8EDF5", bold: true, size: headerSize }),
                wordCell("Semester I", { fill: "E8EDF5", bold: true, size: headerSize }),
                wordCell("Semester II", { fill: "E8EDF5", bold: true, size: headerSize })
            ]
        }),
        ...getAcademicSummaryRowsData().map((item) => new docx.TableRow({
            children: [
                wordCell(item.label, { size: bodySize, align: "left" }),
                wordCell(item.semester1, { size: bodySize }),
                wordCell(item.semester2, { size: bodySize })
            ]
        }))
    ];
    return new docx.Table({
        width: { size: 100, type: docx.WidthType.PERCENTAGE },
        rows
    });
}

function dashboardMonthsToDocxGrid(months) {
    const gridWidthDxa = 15800;
    const monthCellWidthDxa = Math.floor(gridWidthDxa / 4);
    const rows = [];
    for (let r = 0; r < Math.ceil(months.length / 4); r++) {
        const cells = [];
        for (let c = 0; c < 4; c++) {
            const item = months[r * 4 + c];
            cells.push(new docx.TableCell({
                width: { size: monthCellWidthDxa, type: docx.WidthType.DXA },
                borders: noWordBorders(),
                margins: { top: 18, bottom: 18, left: 18, right: 18 },
                children: item ? [dashboardMonthToDocxTable(item.y, item.m, monthCellWidthDxa - 80)] : [new docx.Paragraph("")]
            }));
        }
        rows.push(new docx.TableRow({ children: cells }));
    }
    return new docx.Table({
        width: { size: gridWidthDxa, type: docx.WidthType.DXA },
        borders: noWordBorders(),
        rows
    });
}

function dashboardMonthToDocxTable(year, month, widthDxa = 3800) {
    const dayWidth = Math.floor(widthDxa / 7);
    const compactMargins = { top: 10, bottom: 10, left: 8, right: 8 };
    const rows = [];
    rows.push(new docx.TableRow({
        children: [
            wordCell(String(month + 1), { fill: "93C5FD", bold: true, size: 22, widthDxa: dayWidth, margins: compactMargins }),
            wordCell(monthNames[month], { fill: "93C5FD", bold: true, size: 22, colspan: 5, widthDxa: dayWidth * 5, margins: compactMargins }),
            wordCell(String(year), { fill: "93C5FD", bold: false, size: 22, widthDxa: dayWidth, margins: compactMargins })
        ]
    }));
    rows.push(new docx.TableRow({
        children: dayNames.map((day, index) => wordCell(day, {
            fill: "93C5FD",
            bold: true,
            color: index === 0 || (Number(settings.workDays) === 5 && index === 6) ? "FF0000" : "111111",
            size: 22,
            widthDxa: dayWidth,
            margins: compactMargins
        }))
    }));
    const first = new Date(year, month, 1);
    const start = new Date(year, month, 1 - first.getDay());
    for (let week = 0; week < 6; week++) {
        rows.push(new docx.TableRow({
            children: Array.from({ length: 7 }, (_, day) => {
                const date = new Date(start);
                date.setDate(start.getDate() + week * 7 + day);
                const other = date.getMonth() !== month;
                const events = other ? [] : getDayEvents(formatDate(date));
                const category = events.length ? (categories[events[0].type] || categories.lainnya) : null;
                return wordCell(other ? "" : String(date.getDate()), {
                    fill: category ? normalizeColor(category.color).replace("#", "") : "FFFFFF",
                    color: category ? (category.textColor === "black" ? "111111" : "FFFFFF") : !isWorkingDay(date) ? "FF0000" : "111111",
                    bold: true,
                    size: 22,
                    widthDxa: dayWidth,
                    margins: compactMargins
                });
            })
        }));
    }
    return new docx.Table({
        width: { size: widthDxa, type: docx.WidthType.DXA },
        rows
    });
}

function wordCell(text, options = {}) {
    return new docx.TableCell({
        columnSpan: options.colspan,
        width: options.widthDxa ? { size: options.widthDxa, type: docx.WidthType.DXA } : undefined,
        shading: options.fill && options.fill !== "FFFFFF" ? { fill: options.fill } : undefined,
        verticalAlign: docx.VerticalAlign.CENTER,
        margins: options.margins || { top: 20, bottom: 20, left: 25, right: 25 },
        borders: wordBorders(),
        children: [new docx.Paragraph({
            alignment: options.align === "left" ? docx.AlignmentType.LEFT : docx.AlignmentType.CENTER,
            children: [new docx.TextRun({
                text,
                bold: Boolean(options.bold),
                color: options.color || "111111",
                size: options.size || 14
            })]
        })]
    });
}

function htmlStringToElement(html) {
    const template = document.createElement("template");
    template.innerHTML = html.trim();
    return template.content.firstElementChild;
}

function loadExportOptions() {
    try {
        return JSON.parse(localStorage.getItem(EXPORT_OPTIONS_KEY) || "{}");
    } catch (error) {
        return {};
    }
}

function getSavedExportOptions(kind, target = "") {
    const all = loadExportOptions();
    if (target && all[kind] && all[kind][target]) return all[kind][target];
    return all[kind] && !all[kind][target] && all[kind].paper ? all[kind] : {};
}

function saveExportOptions(kind, options, target = "") {
    try {
        const all = loadExportOptions();
        if (!all[kind] || all[kind].paper) all[kind] = {};
        all[kind][target || "default"] = options;
        localStorage.setItem(EXPORT_OPTIONS_KEY, JSON.stringify(all));
    } catch (error) {
        console.warn("Pengaturan ekspor tidak dapat disimpan.", error);
    }
}

function docxMarginFromMm(mm) {
    const twips = Math.round(Number(mm || DEFAULT_WORD_OPTIONS.margin) * 56.7);
    return { top: twips, right: twips, bottom: twips, left: twips };
}

function getDocxPageSize(portrait) {
    // docx flips width/height internally when orientation is landscape.
    return portrait
        ? { orientation: docx.PageOrientation.PORTRAIT, width: 11906, height: 16838 }
        : { orientation: docx.PageOrientation.LANDSCAPE, width: 11906, height: 16838 };
}

function pageToDocxChildren(page) {
    const children = [];
    $$(":scope > *", page).forEach((node) => appendNodeToDocx(children, node));
    if (!children.length) children.push(new docx.Paragraph(""));
    return children;
}

function appendNodeToDocx(children, node) {
    if (node.classList?.contains("no-export")) return;
    const tag = node.tagName ? node.tagName.toLowerCase() : "";
    if (node.matches?.(".report-footer-grid")) {
        const block = node.querySelector(".signature-block");
        children.push(summarySignatureFooterToDocx(block, { wideSummary: Boolean(node.closest(".akademik-sheet")) }));
        children.push(new docx.Paragraph(""));
        return;
    }
    if (node.matches?.(".signature-row")) {
        children.push(signatureRowToDocx(node));
        children.push(new docx.Paragraph(""));
        return;
    }
    if (node.matches?.(".signature-block")) {
        if (node.classList.contains("right")) children.push(singleSignatureBlockToDocx(node));
        else children.push(...signatureBlockToDocxParagraphs(node, docx.AlignmentType.CENTER));
        children.push(new docx.Paragraph(""));
        return;
    }
    if (tag === "table") {
        children.push(htmlTableToDocx(node));
        children.push(new docx.Paragraph(""));
        return;
    }
    if (node.matches?.(".table-scroll")) {
        const table = node.querySelector("table");
        if (table) {
            children.push(htmlTableToDocx(table));
            children.push(new docx.Paragraph(""));
        }
        return;
    }
    if (["h1", "h2", "h3", "p"].includes(tag) || node.classList?.contains("effective-title")) {
        const text = cleanText(node.innerText || node.textContent);
        if (text) children.push(new docx.Paragraph({
            alignment: node.closest(".report-title") || node.classList?.contains("effective-title") ? docx.AlignmentType.CENTER : docx.AlignmentType.LEFT,
            spacing: { after: tag === "p" ? 80 : 120 },
            children: [new docx.TextRun({ text, bold: tag !== "p", size: tag === "h2" || node.classList?.contains("effective-title") ? 28 : 22 })]
        }));
        return;
    }
    const kids = $$(":scope > *", node);
    if (kids.length) {
        kids.forEach((child) => appendNodeToDocx(children, child));
        return;
    }
    const text = cleanText(node.innerText || node.textContent);
    if (text) children.push(new docx.Paragraph({ children: [new docx.TextRun(text)] }));
}

function singleSignatureBlockToDocx(block) {
    return new docx.Table({
        width: { size: 100, type: docx.WidthType.PERCENTAGE },
        borders: noWordBorders(),
        rows: [new docx.TableRow({
            children: [
                new docx.TableCell({ borders: noWordBorders(), children: [new docx.Paragraph("")] }),
                new docx.TableCell({
                    borders: noWordBorders(),
                    verticalAlign: docx.VerticalAlign.TOP,
                    margins: { top: 80, bottom: 80, left: 80, right: 80 },
                    children: signatureBlockToDocxParagraphs(block, docx.AlignmentType.CENTER)
                })
            ]
        })]
    });
}

function signatureRowToDocx(rowNode) {
    const blocks = $$(".signature-block", rowNode);
    const cells = [0, 1].map((index) => {
        const block = blocks[index];
        return new docx.TableCell({
            borders: noWordBorders(),
            verticalAlign: docx.VerticalAlign.TOP,
            margins: { top: 80, bottom: 80, left: 80, right: 80 },
            children: block ? signatureBlockToDocxParagraphs(block, docx.AlignmentType.CENTER) : [new docx.Paragraph("")]
        });
    });
    return new docx.Table({
        width: { size: 100, type: docx.WidthType.PERCENTAGE },
        borders: noWordBorders(),
        rows: [new docx.TableRow({ children: cells })]
    });
}

function signatureBlockToDocxParagraphs(block, alignment = docx.AlignmentType.LEFT, options = {}) {
    if (!block) return [new docx.Paragraph("")];
    const size = options.size || 20;
    const paragraphs = [];
    $$(":scope > *", block).forEach((child) => {
        if (child.classList.contains("signature-space")) {
            paragraphs.push(new docx.Paragraph({ spacing: { before: 500, after: 140 } }));
            return;
        }
        const text = cleanText(child.innerText || child.textContent);
        if (!text) return;
        paragraphs.push(new docx.Paragraph({
            alignment,
            spacing: { after: 45 },
            children: [new docx.TextRun({
                text,
                bold: child.classList.contains("signature-name"),
                underline: child.classList.contains("signature-name") ? { type: docx.UnderlineType.SINGLE } : undefined,
                size
            })]
        }));
    });
    return paragraphs.length ? paragraphs : [new docx.Paragraph("")];
}

function htmlTableToDocx(table) {
    const rows = Array.from(table.rows).map((htmlRow) => new docx.TableRow({
        children: Array.from(htmlRow.cells).map((cell, cellIndex) => {
            const style = getComputedStyle(cell);
            const fill = (getCellBackgroundArgb(cell) || "FFFFFFFF").slice(2);
            const color = (getCellFontArgb(cell) || "FF111827").slice(2);
            const colspan = Number(cell.getAttribute("colspan") || 1);
            const rowspan = Number(cell.getAttribute("rowspan") || 1);
            const leftAligned = style.textAlign === "left"
                || (table.classList.contains("rekap-table") && [2, 3].includes(cellIndex))
                || (table.classList.contains("effective-summary") && cellIndex === 1);
            return new docx.TableCell({
                columnSpan: colspan > 1 ? colspan : undefined,
                rowSpan: rowspan > 1 ? rowspan : undefined,
                shading: fill !== "FFFFFF" ? { fill } : undefined,
                verticalAlign: docx.VerticalAlign.CENTER,
                margins: { top: 70, bottom: 70, left: 70, right: 70 },
                borders: table.classList.contains("effective-summary") || table.classList.contains("effective-identity-table") ? noWordBorders() : wordBorders(),
                children: [new docx.Paragraph({
                    alignment: leftAligned ? docx.AlignmentType.LEFT : docx.AlignmentType.CENTER,
                    children: [new docx.TextRun({
                        text: cleanText(cell.innerText || cell.textContent),
                        bold: cell.tagName.toLowerCase() === "th" || Number(style.fontWeight) >= 600,
                        color,
                        size: table.classList.contains("academic-board") ? 16 : 20
                    })]
                })]
            });
        })
    }));
    return new docx.Table({
        width: { size: 100, type: docx.WidthType.PERCENTAGE },
        rows
    });
}

function wordBorders() {
    const border = { style: docx.BorderStyle.SINGLE, size: 6, color: "111111" };
    return { top: border, bottom: border, left: border, right: border };
}

function noWordBorders() {
    const border = { style: docx.BorderStyle.NONE, size: 0, color: "FFFFFF" };
    return { top: border, bottom: border, left: border, right: border };
}

function getReportPages(source) {
    const pages = $$(".report-sheet", source);
    return pages.length ? pages : [source];
}

function cloneReportPages(source, exportOptions = {}) {
    return getReportPages(source).map((page) => {
        let clone;
        if (page.classList.contains("dashboard-sheet")) {
            const forPdf = exportOptions.exportKind === "pdf";
            clone = buildPrintableDashboardPage(page, { hideEventText: forPdf, colorizeEventCells: forPdf, columns: 4 });
        }
        else if (page.classList.contains("akademik-sheet")) clone = buildPrintableAcademicPage(page, exportOptions);
        else clone = page.cloneNode(true);
        clone.classList.add("export-mode", "print-fit-page");
        if (exportOptions.exportKind === "pdf") clone.classList.add("pdf-export-mode");
        if (exportOptions.paper) clone.classList.add(`export-page-${exportOptions.paper}`);
        if (exportOptions.orientation) clone.classList.add(`export-${exportOptions.orientation}`);
        $$(".no-export", clone).forEach((node) => node.remove());
        if (exportOptions.exportKind === "pdf" && (clone.classList.contains("dashboard-sheet") || clone.classList.contains("akademik-sheet"))) {
            addPdfLegendToReport(clone);
        }
        return clone;
    });
}

function buildPrintableDashboardPage(page, { hideEventText = false, colorizeEventCells = false, columns = 4 } = {}) {
    const clone = page.cloneNode(true);
    clone.classList.add("print-dashboard-model-a");
    const grid = clone.querySelector(".dashboard-grid");
    if (grid && columns === 4) grid.classList.add("model-a-year");
    if (colorizeEventCells) colorizeDashboardEventCells(clone);
    if (hideEventText) {
        clone.classList.add("export-no-event-text");
        $$(".event-mini", clone).forEach((node) => node.remove());
    }
    $$(".dashboard-stats-table", clone).forEach((node) => node.remove());
    if (hideEventText || colorizeEventCells) normalizeDashboardExportFooter(clone);
    return clone;
}

function normalizeDashboardExportFooter(clone) {
    const summary = clone.querySelector(".dashboard-export-summary");
    const signature = clone.querySelector(".signature-row");
    if (!summary || !signature) return;
    const footer = document.createElement("div");
    footer.className = "report-footer-grid dashboard-pdf-footer";
    const left = document.createElement("div");
    left.className = "dashboard-footer-summary";
    const right = document.createElement("div");
    right.className = "dashboard-footer-right";
    left.appendChild(summary);
    right.appendChild(signature);
    footer.append(left, right);
    clone.appendChild(footer);
}

function colorizeDashboardEventCells(clone) {
    $$(".month-calendar td[data-date]", clone).forEach((cell) => {
        if (cell.classList.contains("other-month")) return;
        const events = getDayEvents(cell.dataset.date);
        if (!events.length) return;
        cell.classList.add("event-cell");
        const light = events.length === 1 && (categories[events[0].type] || {}).textColor === "black";
        if (light) cell.classList.add("light-text");
        cell.style.cssText = `${cell.getAttribute("style") || ""};${getMultiEventStyle(events)}`;
    });
}

function buildPrintableAcademicPage(page) {
    const clone = page.cloneNode(true);
    clone.classList.add("print-academic-page");
    $$(".legend-list", clone).forEach((node) => node.remove());
    return clone;
}

function addPdfLegendToReport(clone) {
    const legend = document.createElement("div");
    legend.className = "legend-list export-legend-list";
    legend.innerHTML = buildLegend();
    const dashboardSummary = clone.querySelector(".dashboard-footer-summary");
    if (dashboardSummary) {
        dashboardSummary.appendChild(legend);
        return;
    }
    const footer = clone.querySelector(".report-footer-grid");
    if (footer) {
        (footer.firstElementChild || footer).appendChild(legend);
        return;
    }
    const signature = clone.querySelector(".signature-row");
    if (signature) signature.insertAdjacentElement("afterend", legend);
    else clone.appendChild(legend);
}

function shouldUsePortrait(page, elementId) {
    return page.classList.contains("effective-sheet") || elementId === "export-efektif" || elementId === "export-rekap";
}

function renderReportTitle(title, subtitle) {
    return `
        <div class="report-title">
            <h2>${title}</h2>
            <h3>${subtitle}</h3>
        </div>
    `;
}

function renderSignatureBlock(type = "kepala", rightOnly = false, overrideDate = "") {
    const dateText = overrideDate || formatDateIndo(settings.ttdDate);
    const headmasterName = formatNameForSignature(settings.headmasterName);
    const teacherName = formatNameForSignature(settings.teacherName);
    const right = `
        <div class="signature-block">
            <p>${escapeHtml(settings.ttdPlace)}, ${escapeHtml(dateText)}</p>
            <p>${type === "guru" ? "Guru Mata Pelajaran" : `Kepala ${escapeHtml(settings.schoolName)}`}</p>
            <div class="signature-space"></div>
            <p class="signature-name">${escapeHtml(type === "guru" ? teacherName : headmasterName)}</p>
            <p>NIP. ${escapeHtml(type === "guru" ? settings.teacherNip : settings.headmasterNip)}</p>
        </div>
    `;
    if (rightOnly) {
        return `
            <div class="signature-block right">
                <p>${escapeHtml(settings.ttdPlace)}, ${escapeHtml(dateText)}</p>
                <p>Kepala ${escapeHtml(settings.schoolName)}</p>
                <div class="signature-space"></div>
                <p class="signature-name">${escapeHtml(headmasterName)}</p>
                <p>NIP. ${escapeHtml(settings.headmasterNip)}</p>
            </div>
        `;
    }
    if (type === "guru") {
        return `
            <div class="signature-row">
                <div class="signature-block">
                    <p>Mengetahui</p>
                    <p>Kepala Sekolah</p>
                    <div class="signature-space"></div>
                    <p class="signature-name">${escapeHtml(headmasterName)}</p>
                    <p>NIP. ${escapeHtml(settings.headmasterNip)}</p>
                </div>
                ${right}
            </div>
        `;
    }
    return `<div class="signature-row"><div class="signature-block"></div>${right}</div>`;
}

function formatNameForSignature(value) {
    return String(value || "")
        .toLowerCase()
        .split(/\s+/)
        .filter(Boolean)
        .map((word) => {
            if (word.includes(".")) {
                return word
                    .split(".")
                    .map((part) => part ? part.charAt(0).toUpperCase() + part.slice(1) : "")
                    .join(".");
            }
            return word.charAt(0).toUpperCase() + word.slice(1);
        })
        .join(" ")
        .replace(/\bS\.pd\b/gi, "S.Pd")
        .replace(/\bM\.m\.pd\b/gi, "M.M.Pd")
        .replace(/\bM\.pd\b/gi, "M.Pd");
}

function downloadBlob(data, mimeType, filename) {
    const blob = data instanceof Blob ? data : new Blob([data], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
}

function isWorkingDay(date) {
    const day = date.getDay();
    if (day === 0) return false;
    if (Number(settings.workDays) === 5 && day === 6) return false;
    return true;
}

function getCurrentAcademicStartYear() {
    const now = new Date();
    return now.getMonth() >= 6 ? now.getFullYear() : now.getFullYear() - 1;
}

function formatDate(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
}

function parseDate(dateStr) {
    const [year, month, day] = String(dateStr || "").split("-").map(Number);
    return new Date(year, month - 1, day);
}

function formatDateIndo(dateStr) {
    if (!dateStr) return "";
    const date = parseDate(dateStr);
    if (Number.isNaN(date.getTime())) return "";
    return `${date.getDate()} ${monthNames[date.getMonth()]} ${date.getFullYear()}`;
}

function calculateDays(startStr, endStr) {
    const start = parseDate(startStr);
    const end = parseDate(endStr);
    return Math.round((end - start) / 86400000) + 1;
}

function normalizeDateCell(value) {
    if (!value) return "";
    if (value instanceof Date) return formatDate(value);
    if (typeof value === "number") {
        const parsed = XLSX.SSF.parse_date_code(value);
        if (!parsed) return "";
        return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
    const text = String(value).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
    const date = new Date(text);
    if (Number.isNaN(date.getTime())) return "";
    return formatDate(date);
}

function normalizeBoolean(value) {
    if (typeof value === "boolean") return value;
    const text = String(value ?? "").trim().toLowerCase();
    return ["true", "ya", "y", "1", "iya", "benar"].includes(text);
}

function clampNumber(value, min, max, fallback) {
    const number = Number(value);
    if (!Number.isFinite(number)) return fallback;
    return Math.min(max, Math.max(min, number));
}

function normalizeColor(value) {
    const text = String(value || "").trim();
    if (/^#[0-9a-f]{6}$/i.test(text)) return text.toLowerCase();
    return "#3b82f6";
}

function sanitizeCategoryKey(value) {
    return String(value || "")
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9_]+/g, "_")
        .replace(/^_+|_+$/g, "");
}

function createId() {
    return `evt-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
}

function sanitizeFileName(value) {
    return String(value || "dokumen").replace(/[\\/:*?"<>|]+/g, "_").replace(/\s+/g, "_").slice(0, 80);
}

function sanitizeSheetName(value) {
    return String(value || "Sheet").replace(/[\\/?*[\]:]+/g, "_").slice(0, 31);
}

function uniqueExcelSheetName(workbook, rawName) {
    const existing = workbook.worksheets.map((sheet) => sheet.name);
    let name = sanitizeSheetName(rawName);
    let index = 1;
    while (existing.includes(name)) {
        name = sanitizeSheetName(`${rawName}_${index}`);
        index++;
    }
    return name;
}

function getTableMaxColumns(table) {
    return Math.max(...Array.from(table.rows).map((row) => Array.from(row.cells).reduce((sum, cell) => sum + Number(cell.getAttribute("colspan") || 1), 0)), 1);
}

function cssColorToHex(value) {
    const argb = cssColorToArgb(value);
    return argb ? argb.slice(2) : "";
}

function cssColorToArgb(value) {
    if (!value || value === "transparent") return "";
    if (value.startsWith("#")) return `FF${value.replace("#", "").toUpperCase()}`;
    const match = value.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/);
    if (!match) return "";
    const alpha = match[4] === undefined ? 1 : Number(match[4]);
    if (alpha === 0) return "";
    const hex = [match[1], match[2], match[3]].map((part) => Number(part).toString(16).padStart(2, "0")).join("").toUpperCase();
    return `FF${hex}`;
}

function cleanText(value) {
    return String(value || "").replace(/\s+/g, " ").trim();
}

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function toast(message, icon = "success") {
    if (window.Swal) {
        Swal.fire({ toast: true, position: "top-end", icon, title: message, timer: 1800, showConfirmButton: false });
    } else {
        alert(message);
    }
}

async function confirmDialog(title, text) {
    if (!window.Swal) return confirm(`${title}\n${text}`);
    const result = await Swal.fire({
        title,
        text,
        icon: "warning",
        showCancelButton: true,
        confirmButtonText: "Ya",
        cancelButtonText: "Batal",
        confirmButtonColor: "#dc2626"
    });
    return result.isConfirmed;
}
