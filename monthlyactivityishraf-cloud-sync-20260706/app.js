const STORAGE_KEY = "monthlyactivityishraf:v1";
const LAST_TAB_KEY = "monthlyactivityishraf:last-tab:v1";
const CLOUD_SYNC_PATH = "/api/data";
const REPORT_STATUS_PATH = "/api/report-status";
const TEMPLATE_PATH = "assets/template.xlsx";
const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const DRAWING_REL_TYPE = `${OFFICE_REL_NS}/drawing`;
const IMAGE_REL_TYPE = `${OFFICE_REL_NS}/image`;
const WORKSHEET_REL_TYPE = `${OFFICE_REL_NS}/worksheet`;
const WORKSHEET_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const SUMMARY_SHEET_FILE = "xl/worksheets/sheet6.xml";
const SUMMARY_SHEET_NAME = "پوختە";
const SIGNATURE_ASSET_VERSION = "20261009-head";
const INSPECTOR_SIGNATURES = [
  { file: "muzaffar.png", width: 842, height: 534, aliases: ["مظفر حيدر مولود", "مظفر حیدر مولود"] },
  { file: "khalid.png", width: 693, height: 325, aliases: ["خالد ابراهيم رحيم", "خالد إبراهيم رحيم"] },
  { file: "barkhawan.png", width: 774, height: 638, aliases: ["بەرخەوان عثمان امين", "بەرخەوان عوسمان ئەمین"] },
  { file: "dana.png", width: 791, height: 387, aliases: ["دانا احمد محمد", "دانا أحمد محمد"] },
];

const kurdishDays = [
  "یەکشەم",
  "دووشەم",
  "سێشەم",
  "چوارشەم",
  "پێنج شەم",
  "هەینی",
  "شەممە",
];

const monthLabels = [
  "1",
  "2",
  "3",
  "4",
  "5",
  "6",
  "7",
  "8",
  "9",
  "10",
  "11",
  "12",
];

const dailySheetFiles = [
  "xl/worksheets/sheet2.xml",
  "xl/worksheets/sheet3.xml",
  "xl/worksheets/sheet4.xml",
  "xl/worksheets/sheet5.xml",
];

const dailySheetNames = ["س١", "س٢", "س٣", "س٤"];

const defaultData = {
  global: {
    specialty: "کۆمپیوتەر",
    studyYear: "٢٠٢٥-٢٠٢٦",
    activityYear: "2026",
    month: 6,
    head: "خالد ابراهیم",
  },
  inspectors: [
    {
      id: 1,
      name: "مظفر حيدر مولود ",
      series: "100",
      schoolsCount: 5,
      schoolsInCharge: { supervision: 5, externalEval: 0, total: 5 },
      schoolsPlanned: { supervision: 5, externalEval: 0, total: 5 },
      schoolsVisited: { supervision: 5, externalEval: 0, other: 0, total: 5 },
      stats: defaultStats(),
      activities: [],
    },
    {
      id: 2,
      name: "عبدالصمد محمد على",
      series: "101",
      schoolsCount: 4,
      schoolsInCharge: { supervision: 5, externalEval: 15, total: 20 },
      schoolsPlanned: { supervision: 4, externalEval: 3, total: 7 },
      schoolsVisited: { supervision: 4, externalEval: 3, other: 0, total: 7 },
      stats: { ...defaultStats(), committees: 1 },
      activities: [],
    },
    {
      id: 3,
      name: "خالد ابراهيم رحيم ",
      series: "102",
      schoolsCount: 3,
      schoolsInCharge: { supervision: 5, externalEval: 0, total: 5 },
      schoolsPlanned: { supervision: 5, externalEval: 0, total: 5 },
      schoolsVisited: { supervision: 5, externalEval: 0, other: 0, total: 5 },
      stats: { ...defaultStats(), committees: 1 },
      activities: [],
    },
    {
      id: 4,
      name: "بەرخەوان عثمان امين ",
      series: "103",
      schoolsCount: 5,
      schoolsInCharge: { supervision: 5, externalEval: 15, total: 20 },
      schoolsPlanned: { supervision: 4, externalEval: 3, total: 7 },
      schoolsVisited: { supervision: 4, externalEval: 3, other: 0, total: 7 },
      stats: { ...defaultStats(), committees: 1 },
      activities: [],
    },
  ],
};

const statFields = [
  ["meetingsParticipated", "کۆبوونەوەی بەشداریکردووە"],
  ["meetingsCompleted", "کۆبوونەوەی ئەنجامیداوە"],
  ["committees", "لێژنەی بەشداریکردووە"],
  ["research", "لێکۆڵینەوە"],
  ["trainingConducted", "وانەی مەشقی"],
  ["seminarsConducted", "کۆڕ و سیمینار"],
  ["parentMeetings", "کۆبوونەوەی دایک و باوک"],
  ["teacherTrainingAttendance", "وانەی مامۆستای ڕاهێنەر"],
  ["reports", "ڕاپۆرت"],
  ["notes", "یاداشت"],
  ["otherActivities", "چاڵاکی جۆراوجۆر"],
];

let appData = loadData();
const lastVisitedTab = loadLastVisitedTab();
let currentView = lastVisitedTab.view;
let currentInspectorId = lastVisitedTab.inspectorId;
let cloudDirty = false;
let cloudHasRecord = false;
let cloudReadSucceeded = false;
let lastCloudPayload = null;
let cloudSyncInProgress = false;
const CLOUD_SYNC_INTERVAL_MS = 5000;
const reportStatuses = new Map();

function defaultStats() {
  return {
    meetingsParticipated: 0,
    meetingsCompleted: 0,
    committees: 0,
    research: 0,
    trainingConducted: 0,
    seminarsConducted: 0,
    parentMeetings: 0,
    teacherTrainingAttendance: 0,
    reports: 0,
    notes: 0,
    otherActivities: 0,
  };
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function normalizeSettings(value = {}, fallback = defaultData.global) {
  const source = fallback || defaultData.global;
  const settings = { ...source, ...(value || {}) };
  settings.specialty ||= source.specialty || defaultData.global.specialty;
  settings.studyYear ||= source.studyYear || defaultData.global.studyYear;
  settings.activityYear = String(settings.activityYear || source.activityYear || defaultData.global.activityYear);
  settings.month = Number(settings.month) || Number(source.month) || defaultData.global.month;
  settings.head ||= source.head || defaultData.global.head;
  return settings;
}

function getInspectorSettings(inspector) {
  if (!inspector) return normalizeSettings(appData.global, defaultData.global);
  inspector.settings = normalizeSettings(inspector.settings, appData.global);
  return inspector.settings;
}

function getSummarySettings() {
  return appData.inspectors?.[0] ? getInspectorSettings(appData.inspectors[0]) : normalizeSettings(appData.global);
}

function loadData() {
  const saved = localStorage.getItem(STORAGE_KEY);
  let data = saved ? safeJsonParse(saved) : clone(defaultData);
  if (!data || !Array.isArray(data.inspectors) || !data.global) {
    data = clone(defaultData);
  }
  normalizeData(data);
  return data;
}

function loadLastVisitedTab() {
  const saved = safeJsonParse(localStorage.getItem(LAST_TAB_KEY) || "");
  const firstInspectorId = appData.inspectors[0]?.id ?? null;
  if (saved?.view === "summary") return { view: "summary", inspectorId: firstInspectorId };

  const inspectorId = Number(saved?.inspectorId);
  if (saved?.view === "inspector" && appData.inspectors.some((inspector) => inspector.id === inspectorId)) {
    return { view: "inspector", inspectorId };
  }
  return { view: "summary", inspectorId: firstInspectorId };
}

function saveLastVisitedTab() {
  localStorage.setItem(LAST_TAB_KEY, JSON.stringify({
    view: currentView,
    inspectorId: currentView === "inspector" ? currentInspectorId : null,
  }));
}

function safeJsonParse(text) {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

function normalizeData(data) {
  data.global = normalizeSettings(data.global, defaultData.global);
  if (!data.inspectors.length) data.inspectors = clone(defaultData.inspectors);
  const usedIds = new Set();
  let maxId = data.inspectors.reduce((max, inspector) => Math.max(max, numberOrZero(inspector.id)), 0);
  data.inspectors.forEach((inspector, index) => {
    const fallback = defaultData.inspectors[index] || {};
    const inspectorId = numberOrZero(inspector.id);
    if (!inspectorId || usedIds.has(inspectorId)) {
      maxId += 1;
      inspector.id = maxId;
    } else {
      inspector.id = inspectorId;
    }
    usedIds.add(Number(inspector.id));
    inspector.name ||= fallback.name || `سەرپەرشتیاری نوێ ${index + 1}`;
    inspector.series ||= String(100 + index);
    inspector.schoolsCount = numberOrZero(inspector.schoolsCount);
    inspector.schoolsInCharge = normalizeTriple(inspector.schoolsInCharge);
    inspector.schoolsPlanned = normalizeTriple(inspector.schoolsPlanned);
    inspector.schoolsVisited = normalizeVisited(inspector.schoolsVisited);
    inspector.stats = { ...defaultStats(), ...(inspector.stats || {}) };
    inspector.settings = normalizeSettings(inspector.settings, data.global);
    inspector.monthlyActivities = normalizeMonthlyActivities(inspector.monthlyActivities);
    migrateVisibleActivitiesToMonth(inspector, inspector.settings);
    loadInspectorActivitiesForSettings(inspector, inspector.settings);
  });
}

function saveData() {
  persistActivityMonths();
  saveLocalData();
  cloudDirty = true;
}

function saveLocalData() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(appData));
}

async function loadCloudData(options = {}) {
  try {
    const response = await fetch(CLOUD_SYNC_PATH, { cache: "no-store", redirect: "manual" });
    if (!response.ok) return false;
    cloudReadSucceeded = true;

    const payload = await response.text();
    if (cloudDirty) return true;
    if (!payload || payload === "null") {
      cloudHasRecord = false;
      lastCloudPayload = null;
      if (!options.quiet) showToast("هیچ پاشەکەوتێکی هەور نییە؛ داتاکەت هەڵبگرە بۆ بارکردنی.");
      return true;
    }

    if (payload === lastCloudPayload) return true;
    const cloudData = JSON.parse(payload);
    if (!cloudData || !Array.isArray(cloudData.inspectors) || !cloudData.global) return false;
    appData = cloudData;
    normalizeData(appData);
    currentInspectorId = appData.inspectors.some((inspector) => inspector.id === currentInspectorId)
      ? currentInspectorId
      : appData.inspectors[0]?.id ?? null;
    saveLocalData();
    cloudDirty = false;
    cloudHasRecord = true;
    lastCloudPayload = payload;
    render();
    if (!options.quiet) showToast("داتاکان لەسەر هەوری هاوبەش نوێکرانەوە.");
    return true;
  } catch (error) {
    console.warn("Cloud data load failed", error);
    return false;
  }
}

async function saveCloudData() {
  const snapshot = JSON.stringify(appData);
  try {
    const response = await fetch(CLOUD_SYNC_PATH, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: snapshot,
      redirect: "manual",
    });
    if (response.ok) {
      lastCloudPayload = snapshot;
      cloudHasRecord = true;
      if (snapshot === JSON.stringify(appData)) cloudDirty = false;
    }
    return response.ok;
  } catch (error) {
    console.warn("Cloud data save failed", error);
    return false;
  }
}

async function syncCloudData() {
  if (cloudSyncInProgress) return;
  cloudSyncInProgress = true;
  try {
    if (cloudDirty) {
      const saved = await saveCloudData();
      if (!saved) console.warn("Automatic cloud save failed; it will retry.");
      return;
    }
    await loadCloudData({ quiet: true });
  } finally {
    cloudSyncInProgress = false;
  }
}

async function resumeCloudSession(intent) {
  if (intent === "save") {
    const saved = await saveCloudData();
    if (!saved) cloudDirty = true;
    showToast(saved
      ? "زانیارییەکان لەسەر هەوری هاوبەش هەڵگیران."
      : "لەم ئامێرەدا هەڵگیرا، بەڵام هەڵگرتنی هەوری سەرکەوتوو نەبوو.");
    return;
  }

  const loaded = await loadCloudData();
  if (!loaded) {
    showToast("پەیوەندی بە هەورەوە شکستی هێنا؛ دووبارە هەوڵ بدە.");
  }
}

function normalizeTriple(value = {}) {
  const supervision = numberOrZero(value.supervision);
  const externalEval = numberOrZero(value.externalEval);
  return { supervision, externalEval, total: supervision + externalEval };
}

function normalizeVisited(value = {}) {
  const supervision = numberOrZero(value.supervision);
  const externalEval = numberOrZero(value.externalEval);
  const other = numberOrZero(value.other);
  return { supervision, externalEval, other, total: supervision + externalEval + other };
}

function numberOrZero(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function getDaysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

function isWeekend(year, month, day) {
  const jsDay = new Date(year, month - 1, day).getDay();
  return jsDay === 5 || jsDay === 6;
}

function getKurdishDayName(year, month, day) {
  return kurdishDays[new Date(year, month - 1, day).getDay()];
}

function generateMonthActivities(year, month, existingActivities = []) {
  const safeYear = Number(year) || new Date().getFullYear();
  const safeMonth = Number(month) || 1;
  const days = getDaysInMonth(safeYear, safeMonth);
  const result = [];
  for (let day = 1; day <= days; day += 1) {
    const existing = existingActivities[day - 1] || {};
    const weekend = isWeekend(safeYear, safeMonth, day);
    const defaultActivity = weekend ? "پشوو" : "ئامادەبوون لە فەرمانگە";
    const hasSavedActivity = Object.prototype.hasOwnProperty.call(existing, "activity");
    result.push({
      day: getKurdishDayName(safeYear, safeMonth, day),
      date: `${safeYear}-${String(safeMonth).padStart(2, "0")}-${String(day).padStart(2, "0")}`,
      activity: hasSavedActivity ? String(existing.activity ?? "") : defaultActivity,
      isHoliday: existing.isHoliday ?? weekend,
    });
  }
  return result;
}

function getActivityMonthKey(year, month) {
  const safeYear = Number(year) || new Date().getFullYear();
  const safeMonth = Number(month) || 1;
  return `${safeYear}-${String(safeMonth).padStart(2, "0")}`;
}

function getActivityMonthKeyFromSettings(settings) {
  return getActivityMonthKey(settings.activityYear, settings.month);
}

function normalizeMonthlyActivities(value) {
  if (!value || Array.isArray(value) || typeof value !== "object") return {};
  return Object.fromEntries(
    Object.entries(value).map(([key, activities]) => [
      key,
      Array.isArray(activities) ? activities.map(cleanActivityEntry) : [],
    ]),
  );
}

function cleanActivityEntry(activity = {}) {
  return {
    activity: Object.prototype.hasOwnProperty.call(activity, "activity") ? String(activity.activity ?? "") : "",
    isHoliday: Boolean(activity.isHoliday),
  };
}

function serializeActivityList(activities = []) {
  return activities.map(cleanActivityEntry);
}

function migrateVisibleActivitiesToMonth(inspector, settings = getInspectorSettings(inspector)) {
  if (!Array.isArray(inspector.activities) || !inspector.activities.length) return;
  inspector.monthlyActivities = normalizeMonthlyActivities(inspector.monthlyActivities);
  const key = getActivityMonthKeyFromSettings(settings);
  if (!inspector.monthlyActivities[key]?.length) {
    inspector.monthlyActivities[key] = serializeActivityList(inspector.activities);
  }
}

function saveInspectorActivitiesForSettings(inspector, settings = getInspectorSettings(inspector)) {
  if (!inspector) return;
  inspector.monthlyActivities = normalizeMonthlyActivities(inspector.monthlyActivities);
  inspector.monthlyActivities[getActivityMonthKeyFromSettings(settings)] = serializeActivityList(inspector.activities || []);
}

function loadInspectorActivitiesForSettings(inspector, settings = getInspectorSettings(inspector)) {
  inspector.monthlyActivities = normalizeMonthlyActivities(inspector.monthlyActivities);
  const year = Number(settings.activityYear) || new Date().getFullYear();
  const month = Number(settings.month) || 1;
  const key = getActivityMonthKey(year, month);
  inspector.activities = generateMonthActivities(year, month, inspector.monthlyActivities[key] || []);
  inspector.monthlyActivities[key] = serializeActivityList(inspector.activities);
}

function persistActivityMonths() {
  appData.inspectors.forEach((inspector) => saveInspectorActivitiesForSettings(inspector));
}

function regenerateInspectorActivities(inspector) {
  const settings = getInspectorSettings(inspector);
  loadInspectorActivitiesForSettings(inspector, settings);
}

function regenerateActivities() {
  appData.inspectors.forEach((inspector) => {
    regenerateInspectorActivities(inspector);
  });
}

async function init() {
  if (window.location.pathname.replace(/\/$/, "") === "/boss") {
    await initBossPortal();
    return;
  }
  bindEvents();
  render();
  const hadLocalData = Boolean(localStorage.getItem(STORAGE_KEY));
  const cloudIntent = new URLSearchParams(window.location.search).get("cloud");
  if (cloudIntent === "save" || cloudIntent === "load") {
    window.history.replaceState({}, "", window.location.pathname);
    await resumeCloudSession(cloudIntent);
  } else {
    await loadCloudData({ quiet: true });
  }
  if (cloudReadSucceeded && !cloudHasRecord && hadLocalData) cloudDirty = true;
  await refreshInspectorReportStatus();
  window.setInterval(syncCloudData, CLOUD_SYNC_INTERVAL_MS);
  window.setInterval(refreshInspectorReportStatus, CLOUD_SYNC_INTERVAL_MS);
}

function renderMonthOptions(select, selectedValue) {
  select.innerHTML = monthLabels
    .map((label, index) => `<option value="${index + 1}">${label}</option>`)
    .join("");
  select.value = String(selectedValue);
}

function bindEvents() {
  document.getElementById("excelButton").addEventListener("click", generateWorkbook);
  document.getElementById("printButton").addEventListener("click", printAllSheets);
  document.getElementById("printSummaryButton").addEventListener("click", printSummarySheet);
  document.getElementById("addInspectorButton").addEventListener("click", addInspector);
  document.getElementById("mainSaveButton").addEventListener("click", saveMainPage);
  document.getElementById("mainSheet").addEventListener("input", (event) => updateMainPage(event, false));
  document.getElementById("mainSheet").addEventListener("change", (event) => updateMainPage(event, true));
  document.getElementById("mainSheet").addEventListener("click", handleMainPageClick);
  document.getElementById("inspectorStudyYearInput").addEventListener("input", updateInspectorSettings);
  document.getElementById("inspectorStudyYearInput").addEventListener("change", updateInspectorSettings);
  document.getElementById("inspectorActivityYearInput").addEventListener("change", updateInspectorSettings);
  document.getElementById("inspectorMonthInput").addEventListener("change", updateInspectorSettings);
  document.getElementById("inspectorSaveButton").addEventListener("click", saveCurrentInspector);
  document.getElementById("inspectorPrintButton").addEventListener("click", printCurrentInspector);
  document.getElementById("inspectorSendButton").addEventListener("click", sendCurrentReport);
  document.getElementById("inspectorNameInput").addEventListener("input", () => updateInspectorBasics(false));
  document.getElementById("inspectorNameInput").addEventListener("change", () => updateInspectorBasics(true));
  document.getElementById("inspectorSeriesInput").addEventListener("input", () => updateInspectorBasics(false));
  document.getElementById("inspectorSeriesInput").addEventListener("change", () => updateInspectorBasics(true));
  document.getElementById("inspectorSchoolsInput").addEventListener("input", () => updateInspectorBasics(false));
  document.getElementById("inspectorSchoolsInput").addEventListener("change", () => updateInspectorBasics(true));
  document.getElementById("schoolCounts").addEventListener("input", (event) => updateSchoolCount(event, false));
  document.getElementById("schoolCounts").addEventListener("change", (event) => updateSchoolCount(event, true));
  document.getElementById("statCounts").addEventListener("input", (event) => updateStatCount(event, false));
  document.getElementById("statCounts").addEventListener("change", (event) => updateStatCount(event, true));
  document.getElementById("activitiesBody").addEventListener("input", updateActivity);
  document.getElementById("activitiesBody").addEventListener("change", updateActivity);
  window.addEventListener("beforeunload", () => {
    persistVisibleEdits();
    saveData();
  });
}

function render() {
  if (currentView === "inspector" && !appData.inspectors.some((inspector) => inspector.id === currentInspectorId)) {
    currentView = "summary";
    currentInspectorId = appData.inspectors[0]?.id ?? null;
  }
  saveLastVisitedTab();
  renderTabs();
  if (currentView === "summary") {
    document.getElementById("summaryPanel").hidden = false;
    document.getElementById("inspectorPanel").hidden = true;
    renderSummary();
  } else {
    document.getElementById("summaryPanel").hidden = true;
    document.getElementById("inspectorPanel").hidden = false;
    renderInspector();
  }
}

function renderTabs() {
  const tabs = document.getElementById("tabs");
  const summaryActive = currentView === "summary" ? " active" : "";
  const inspectorTabs = appData.inspectors
    .map((inspector) => {
      const active = currentView === "inspector" && currentInspectorId === inspector.id ? " active" : "";
      return `<button type="button" class="tab${active}" data-view="inspector" data-id="${inspector.id}">${escapeHtml(inspector.name.trim())}</button>`;
    })
    .join("");
  tabs.innerHTML = `<button type="button" class="tab${summaryActive}" data-view="summary">سەرەکی</button>${inspectorTabs}`;
  tabs.querySelectorAll(".tab").forEach((button) => {
    button.addEventListener("click", () => {
      currentView = button.dataset.view;
      if (currentView === "inspector") currentInspectorId = Number(button.dataset.id);
      render();
    });
  });
}

function renderSummary() {
  const summarySettings = getSummarySettings();
  document.getElementById("mainStudyYearInput").value = summarySettings.studyYear;
  document.getElementById("mainHeadInput").value = summarySettings.head;
  document.getElementById("mainInspectorsBody").innerHTML = appData.inspectors
    .map((inspector, index) => {
      return `<tr>
        <td>${index + 1}</td>
        <td><input class="main-name-input" type="text" value="${escapeHtml(inspector.name)}" data-main-field="name" data-index="${index}" autocomplete="off"></td>
        <td><input type="text" value="${escapeHtml(inspector.series)}" data-main-field="series" data-index="${index}" autocomplete="off"></td>
        <td><input type="number" min="0" value="${inspector.schoolsCount}" data-main-field="schoolsCount" data-index="${index}"></td>
        <td><button type="button" class="button main-remove-button" data-remove-inspector="${index}">سڕینەوە</button></td>
      </tr>`;
    })
    .join("");
}

function renderInspector() {
  const inspector = getCurrentInspector();
  if (!inspector) {
    currentView = "summary";
    render();
    return;
  }
  const settings = getInspectorSettings(inspector);
  document.getElementById("inspectorTitle").textContent = inspector.name;
  renderInspectorSendButton(inspector);
  document.getElementById("inspectorMeta").textContent = `مانگی ${settings.month} / ${settings.studyYear}`;
  document.getElementById("inspectorStudyYearInput").value = settings.studyYear;
  document.getElementById("inspectorActivityYearInput").value = settings.activityYear;
  renderMonthOptions(document.getElementById("inspectorMonthInput"), settings.month);
  document.getElementById("inspectorNameInput").value = inspector.name;
  document.getElementById("inspectorSeriesInput").value = inspector.series;
  document.getElementById("inspectorSchoolsInput").value = inspector.schoolsCount;
  renderSchoolCounts(inspector);
  renderStatCounts(inspector);
  renderActivities(inspector);
}

function reportStatusKey(inspectorId, year, month) {
  return `${Number(inspectorId)}:${Number(year)}:${Number(month)}`;
}

function isReportSent(inspectorId, year, month) {
  return reportStatuses.has(reportStatusKey(inspectorId, year, month));
}

function renderInspectorSendButton(inspector) {
  const button = document.getElementById("inspectorSendButton");
  if (!button) return;
  const settings = getInspectorSettings(inspector);
  const sent = isReportSent(inspector.id, settings.activityYear, settings.month);
  button.classList.toggle("sent", sent);
  button.textContent = sent ? "نێردرا" : "ناردن";
  button.setAttribute("aria-label", sent ? "نێردرا" : "ناردنی ڕاپۆرت بۆ سەرپەرشتیار");
}

async function fetchReportStatuses(year, month) {
  const response = await fetch(`${REPORT_STATUS_PATH}?year=${encodeURIComponent(year)}&month=${encodeURIComponent(month)}`, { cache: "no-store" });
  if (!response.ok) throw new Error("Could not load report statuses");
  const data = await response.json();
  return Array.isArray(data.reports) ? data.reports : [];
}

async function refreshInspectorReportStatus() {
  const inspector = getCurrentInspector();
  if (!inspector) return;
  const settings = getInspectorSettings(inspector);
  try {
    const reports = await fetchReportStatuses(settings.activityYear, settings.month);
    for (const [key, value] of reportStatuses) {
      if (key.endsWith(`:${Number(settings.activityYear)}:${Number(settings.month)}`)) reportStatuses.delete(key);
    }
    reports.forEach((report) => reportStatuses.set(
      reportStatusKey(report.inspectorId, settings.activityYear, settings.month), report.sentAt,
    ));
    renderInspectorSendButton(inspector);
  } catch (error) {
    console.warn("Report status refresh failed", error);
  }
}

async function sendCurrentReport() {
  const inspector = getCurrentInspector();
  if (!inspector) return;
  persistVisibleEdits();
  readInspectorSettingsFromFields(inspector);
  normalizeData(appData);
  saveData();
  const saved = await saveCloudData();
  if (!saved) {
    showToast("سەرەتا پاشەکەوتی هەوری سەرکەوتوو نەبوو؛ ناردن ئەنجام نەدرا.");
    return;
  }
  const settings = getInspectorSettings(inspector);
  try {
    const response = await fetch(`${REPORT_STATUS_PATH}/send`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ inspectorId: inspector.id, year: settings.activityYear, month: settings.month }),
    });
    if (!response.ok) throw new Error("Could not send report");
    await refreshInspectorReportStatus();
    showToast("ڕاپۆرتەکە نێردرا بۆ چاپکردن.");
  } catch (error) {
    console.error("Report send failed", error);
    showToast("ناردنی ڕاپۆرت سەرکەوتوو نەبوو؛ دووبارە هەوڵ بدە.");
  }
}

async function initBossPortal() {
  document.body.classList.add("boss-mode");
  const today = new Date();
  const period = { activityYear: String(today.getFullYear()), month: today.getMonth() + 1 };
  const boss = document.createElement("main");
  boss.id = "bossPortal";
  boss.className = "boss-portal";
  boss.innerHTML = `
    <header class="boss-header">
      <div><h1>کاری مانگانەی سەرپەرشتیاران</h1><p>ڕاپۆرتە نێردراوەکان ئامادەن بۆ چاپکردن.</p></div>
      <div class="boss-actions">
        <label>ساڵ <input id="bossYear" type="number" min="2000" max="2100" value="${escapeHtml(period.activityYear)}"></label>
        <label>مانگ <select id="bossMonth"></select></label>
        <button class="button ghost" id="bossRefresh" type="button">نوێکردنەوە</button>
        <button class="button print" id="bossPrintAll" type="button" disabled>چاپکردنی هەموو نێردراوەکان</button>
        <button class="button print" id="bossPrintSummary" type="button" disabled>چاپکردنی پوختە</button>
        <button class="button primary" id="bossExcelExport" type="button">دابەزاندنی Excel</button>
      </div>
    </header>
    <section class="boss-list-panel"><div class="boss-list-heading"><h2>سەرپەرشتیاران</h2><span id="bossPeriod"></span></div><div id="bossInspectorList" class="boss-inspector-list" aria-live="polite"></div></section>`;
  document.body.insertBefore(boss, document.body.firstChild);
  const monthSelect = boss.querySelector("#bossMonth");
  renderMonthOptions(monthSelect, period.month);
  boss.querySelector("#bossRefresh").addEventListener("click", refreshBossPortal);
  boss.querySelector("#bossYear").addEventListener("change", refreshBossPortal);
  monthSelect.addEventListener("change", refreshBossPortal);
  boss.querySelector("#bossPrintAll").addEventListener("click", () => printBossReports());
  boss.querySelector("#bossPrintSummary").addEventListener("click", printBossSummary);
  boss.querySelector("#bossExcelExport").addEventListener("click", exportBossWorkbook);
  await refreshBossPortal();
  window.setInterval(() => {
    const now = new Date();
    const yearInput = document.getElementById("bossYear");
    const monthInput = document.getElementById("bossMonth");
    if (!yearInput || !monthInput) return;
    const year = String(now.getFullYear());
    const month = String(now.getMonth() + 1);
    if (yearInput.value !== year || monthInput.value !== month) {
      yearInput.value = year;
      monthInput.value = month;
      refreshBossPortal();
    }
  }, 60000);
}

function getBossPeriod() {
  const year = Number(document.getElementById("bossYear").value);
  const month = Number(document.getElementById("bossMonth").value);
  if (!Number.isInteger(year) || year < 2000 || year > 2100 || !Number.isInteger(month) || month < 1 || month > 12) {
    throw new Error("Invalid report period");
  }
  return { year, month };
}

async function loadBossData() {
  const response = await fetch(CLOUD_SYNC_PATH, { cache: "no-store" });
  if (!response.ok) throw new Error("Could not load shared data");
  const payload = await response.json();
  if (!payload || !Array.isArray(payload.inspectors) || !payload.global) throw new Error("No shared data is available");
  normalizeData(payload);
  return payload;
}

async function refreshBossPortal() {
  const list = document.getElementById("bossInspectorList");
  if (!list) return;
  const { year, month } = getBossPeriod();
  document.getElementById("bossPeriod").textContent = `${month} / ${year}`;
  list.innerHTML = `<p class="boss-message">داتا بار دەکرێت…</p>`;
  try {
    const [data, reports] = await Promise.all([loadBossData(), fetchReportStatuses(year, month)]);
    const sent = new Set(reports.map((report) => Number(report.inspectorId)));
    list.innerHTML = data.inspectors.map((inspector) => `
      <div class="boss-inspector-row">
        <span class="boss-inspector-name">${escapeHtml(inspector.name.trim())}</span>
        <button type="button" class="button boss-print-one${sent.has(Number(inspector.id)) ? " ready" : ""}" data-boss-print="${Number(inspector.id)}" ${sent.has(Number(inspector.id)) ? "" : "disabled"}>
          ${sent.has(Number(inspector.id)) ? "ئامادەی چاپ" : "نێردراو نییە"}
        </button>
      </div>`).join("") || `<p class="boss-message">هیچ سەرپەرشتیارێک نییە.</p>`;
    list.querySelectorAll("[data-boss-print]").forEach((button) => button.addEventListener("click", () => {
      printBossReports([Number(button.dataset.bossPrint)]);
    }));
    const printAll = document.getElementById("bossPrintAll");
    printAll.disabled = sent.size === 0;
    printAll.classList.toggle("ready", sent.size > 0);
    const everyInspectorSent = data.inspectors.length > 0 && data.inspectors.every((inspector) => sent.has(Number(inspector.id)));
    const printSummary = document.getElementById("bossPrintSummary");
    printSummary.disabled = !everyInspectorSent;
    printSummary.classList.toggle("ready", everyInspectorSent);
  } catch (error) {
    console.error("Boss portal refresh failed", error);
    list.innerHTML = `<p class="boss-message error">پەیوەندی بە داتای هاوبەش سەرکەوتوو نەبوو. دووبارە نوێی بکەرەوە.</p>`;
    document.getElementById("bossPrintAll").disabled = true;
    document.getElementById("bossPrintSummary").disabled = true;
  }
}

async function printBossSummary() {
  try {
    const { year, month } = getBossPeriod();
    const [data, reports] = await Promise.all([loadBossData(), fetchReportStatuses(year, month)]);
    const sentIds = new Set(reports.map((report) => Number(report.inspectorId)));
    if (!data.inspectors.length || !data.inspectors.every((inspector) => sentIds.has(Number(inspector.id)))) {
      await refreshBossPortal();
      return;
    }
    data.inspectors.forEach((inspector) => {
      inspector.settings = normalizeSettings({ ...getInspectorSettings(inspector), activityYear: String(year), month }, data.global);
      loadInspectorActivitiesForSettings(inspector, inspector.settings);
    });
    appData = data;
    const printRoot = document.getElementById("printRoot");
    printRoot.innerHTML = buildSummaryPrintPage();
    await printAndWaitForClose();
  } catch (error) {
    console.error("Boss summary print failed", error);
    const list = document.getElementById("bossInspectorList");
    if (list) list.insertAdjacentHTML("afterbegin", `<p class="boss-message error">چاپی پوختە سەرکەوتوو نەبوو؛ دووبارە هەوڵ بدە.</p>`);
  }
}

async function exportBossWorkbook() {
  try {
    const { year, month } = getBossPeriod();
    const data = await loadBossData();
    data.inspectors.forEach((inspector) => {
      inspector.settings = normalizeSettings({
        ...getInspectorSettings(inspector),
        activityYear: String(year),
        month,
      }, data.global);
      loadInspectorActivitiesForSettings(inspector, inspector.settings);
    });
    data.global.activityYear = String(year);
    data.global.month = month;
    appData = data;
    await generateWorkbook({ persist: false });
  } catch (error) {
    console.error("Boss Excel export failed", error);
    const list = document.getElementById("bossInspectorList");
    if (list) list.insertAdjacentHTML("afterbegin", `<p class="boss-message error">دروستکردنی فایلی Excel سەرکەوتوو نەبوو؛ دووبارە هەوڵ بدە.</p>`);
  }
}

async function printBossReports(onlyInspectorIds = null) {
  const { year, month } = getBossPeriod();
  try {
    const [data, reports] = await Promise.all([loadBossData(), fetchReportStatuses(year, month)]);
    const sentIds = new Set(reports.map((report) => Number(report.inspectorId)));
    const chosenIds = onlyInspectorIds ? onlyInspectorIds.filter((id) => sentIds.has(id)) : [...sentIds];
    if (!chosenIds.length) return;
    const pages = chosenIds.flatMap((id) => {
      const index = data.inspectors.findIndex((inspector) => Number(inspector.id) === id);
      if (index < 0) return [];
      const inspector = data.inspectors[index];
      inspector.settings = normalizeSettings({ ...getInspectorSettings(inspector), activityYear: String(year), month }, data.global);
      loadInspectorActivitiesForSettings(inspector, inspector.settings);
      return [buildDailyPrintPage(inspector, index)];
    });
    if (!pages.length) return;
    document.getElementById("printRoot").innerHTML = pages.join("");
    await printAndWaitForClose();
    const printResponses = await Promise.all(chosenIds.map((inspectorId) => fetch(`${REPORT_STATUS_PATH}/printed`, {
      method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({ inspectorId, year, month }),
    })));
    if (printResponses.some((response) => !response.ok)) throw new Error("Could not update printed status");
    await refreshBossPortal();
  } catch (error) {
    console.error("Boss print failed", error);
    const list = document.getElementById("bossInspectorList");
    if (list) list.insertAdjacentHTML("afterbegin", `<p class="boss-message error">چاپکردن سەرکەوتوو نەبوو؛ دڵنیابە داتاکان بارکراون و دووبارە هەوڵ بدە.</p>`);
  }
}

async function printAndWaitForClose() {
  await waitForPrintImages();
  return new Promise((resolve) => {
    let settled = false;
    const finish = () => {
      if (settled) return;
      settled = true;
      window.removeEventListener("afterprint", finish);
      resolve();
    };
    window.addEventListener("afterprint", finish, { once: true });
    window.print();
    window.setTimeout(finish, 120000);
  });
}

function waitForPrintImages() {
  const images = Array.from(document.getElementById("printRoot")?.querySelectorAll("img") || []);
  return Promise.all(images.map((image) => {
    if (image.complete && image.naturalWidth) return Promise.resolve();
    if (typeof image.decode === "function") return image.decode().catch(() => {});
    return new Promise((resolve) => {
      image.addEventListener("load", resolve, { once: true });
      image.addEventListener("error", resolve, { once: true });
    });
  }));
}

function renderSchoolCounts(inspector) {
  const fields = [
    ["schoolsInCharge", "supervision", "لە ئەستۆ / سەرپەرشتیاری"],
    ["schoolsInCharge", "externalEval", "لە ئەستۆ / دەرەکی"],
    ["schoolsPlanned", "supervision", "پلاندان / سەرپەرشتیاری"],
    ["schoolsPlanned", "externalEval", "پلاندان / دەرەکی"],
    ["schoolsVisited", "supervision", "سەردان / سەرپەرشتیاری"],
    ["schoolsVisited", "externalEval", "سەردان / دەرەکی"],
    ["schoolsVisited", "other", "سەردانی تر"],
  ];
  document.getElementById("schoolCounts").innerHTML = fields
    .map(([group, field, label]) => {
      const value = inspector[group]?.[field] ?? 0;
      return `<label>
        <span>${label}</span>
        <input type="number" min="0" value="${value}" data-group="${group}" data-field="${field}">
      </label>`;
    })
    .join("");
}

function renderStatCounts(inspector) {
  document.getElementById("statCounts").innerHTML = statFields
    .map(([field, label]) => `<label>
      <span>${label}</span>
      <input type="number" min="0" value="${inspector.stats[field] || 0}" data-field="${field}">
    </label>`)
    .join("");
}

function renderActivities(inspector) {
  document.getElementById("activitiesBody").innerHTML = inspector.activities
    .map((activity, index) => {
      const weekend = activity.isHoliday ? " class=\"weekend\"" : "";
      return `<tr${weekend}>
        <td>${escapeHtml(activity.day)}</td>
        <td>${escapeHtml(formatDisplayDate(activity.date))}</td>
        <td><input type="text" list="activityPresets" value="${escapeHtml(activity.activity)}" data-index="${index}"></td>
      </tr>`;
    })
    .join("");
}

function updateInspectorBasics(shouldRender = true) {
  const inspector = getCurrentInspector();
  if (!inspector) return;
  inspector.name = document.getElementById("inspectorNameInput").value;
  inspector.series = document.getElementById("inspectorSeriesInput").value;
  inspector.schoolsCount = numberOrZero(document.getElementById("inspectorSchoolsInput").value);
  saveData();
  if (shouldRender) render();
}

function persistVisibleActivityInputs(inspector, settings = getInspectorSettings(inspector)) {
  const inspectorPanel = document.getElementById("inspectorPanel");
  if (!inspector || !inspectorPanel || inspectorPanel.hidden || inspector.id !== currentInspectorId) return;
  document.querySelectorAll("#activitiesBody input[data-index]").forEach((input) => {
    const index = Number(input.dataset.index);
    if (!inspector.activities[index]) return;
    inspector.activities[index].activity = input.value;
    inspector.activities[index].isHoliday = input.value.trim() === "پشوو";
  });
  saveInspectorActivitiesForSettings(inspector, settings);
}

function readInspectorSettingsFromFields(inspector) {
  const previous = getInspectorSettings(inspector);
  persistVisibleActivityInputs(inspector, previous);
  const next = normalizeSettings(
    {
      specialty: previous.specialty,
      studyYear: document.getElementById("inspectorStudyYearInput").value.trim(),
      activityYear: document.getElementById("inspectorActivityYearInput").value || String(new Date().getFullYear()),
      month: Number(document.getElementById("inspectorMonthInput").value),
      head: appData.global.head || previous.head,
    },
    previous,
  );
  const changedMonth =
    Number(previous.activityYear) !== Number(next.activityYear) || Number(previous.month) !== Number(next.month);
  if (changedMonth) saveInspectorActivitiesForSettings(inspector, previous);
  inspector.settings = next;
  if (changedMonth) loadInspectorActivitiesForSettings(inspector, next);
  else saveInspectorActivitiesForSettings(inspector, next);
}

function updateInspectorSettings(event = null) {
  const inspector = getCurrentInspector();
  if (!inspector) return;
  readInspectorSettingsFromFields(inspector);
  saveData();
  if (event?.type !== "input") render();
}

async function saveCurrentInspector() {
  const inspector = getCurrentInspector();
  if (!inspector) return;
  readInspectorSettingsFromFields(inspector);
  normalizeData(appData);
  saveData();
  render();
  const saved = await saveCloudData();
  showToast(saved
    ? "زانیارییەکان لەم ئامێرە و لەسەر هەوری هاوبەش هەڵگیران."
    : "لەم ئامێرەدا هەڵگیرا، بەڵام پاشەکەوتی هەور سەرکەوتوو نەبوو.");
}

async function saveMainPage() {
  persistVisibleEdits();
  normalizeData(appData);
  saveData();
  render();
  const saved = await saveCloudData();
  showToast(saved
    ? "زانیارییەکان لەم ئامێرە و لەسەر هەوری هاوبەش هەڵگیران."
    : "لەم ئامێرەدا هەڵگیرا، بەڵام پاشەکەوتی هەور سەرکەوتوو نەبوو.");
}

function persistVisibleEdits() {
  const summaryPanel = document.getElementById("summaryPanel");
  if (summaryPanel && !summaryPanel.hidden) {
    const studyYear = document.getElementById("mainStudyYearInput")?.value.trim();
    const head = document.getElementById("mainHeadInput")?.value.trim();
    appData.global = normalizeSettings({ ...appData.global, studyYear, head }, appData.global);
    appData.inspectors.forEach((inspector) => {
      const previous = getInspectorSettings(inspector);
      inspector.settings = normalizeSettings({ ...previous, studyYear, head }, previous);
    });
    document.querySelectorAll("#mainInspectorsBody [data-main-field]").forEach((input) => {
      const inspector = appData.inspectors[Number(input.dataset.index)];
      if (!inspector) return;
      if (input.dataset.mainField === "name") inspector.name = input.value;
      if (input.dataset.mainField === "series") inspector.series = input.value;
      if (input.dataset.mainField === "schoolsCount") inspector.schoolsCount = numberOrZero(input.value);
    });
  }

  const inspectorPanel = document.getElementById("inspectorPanel");
  if (inspectorPanel && !inspectorPanel.hidden) {
    const inspector = getCurrentInspector();
    if (!inspector) return;
    inspector.name = document.getElementById("inspectorNameInput")?.value ?? inspector.name;
    inspector.series = document.getElementById("inspectorSeriesInput")?.value ?? inspector.series;
    inspector.schoolsCount = numberOrZero(document.getElementById("inspectorSchoolsInput")?.value ?? inspector.schoolsCount);
    persistVisibleActivityInputs(inspector);
    readInspectorSettingsFromFields(inspector);
    document.querySelectorAll("#schoolCounts input").forEach((input) => {
      const group = input.dataset.group;
      const field = input.dataset.field;
      if (inspector[group] && field) inspector[group][field] = numberOrZero(input.value);
    });
    inspector.schoolsInCharge = normalizeTriple(inspector.schoolsInCharge);
    inspector.schoolsPlanned = normalizeTriple(inspector.schoolsPlanned);
    inspector.schoolsVisited = normalizeVisited(inspector.schoolsVisited);
    document.querySelectorAll("#statCounts input").forEach((input) => {
      if (input.dataset.field) inspector.stats[input.dataset.field] = numberOrZero(input.value);
    });
  }
}

function updateMainPage(event, shouldRender = true) {
  const target = event.target;
  const isMainSetting =
    target.id === "mainStudyYearInput" ||
    target.id === "mainHeadInput";
  if (isMainSetting) {
    const next = {
      studyYear: document.getElementById("mainStudyYearInput").value.trim(),
      head: document.getElementById("mainHeadInput").value.trim(),
    };
    appData.global = normalizeSettings({ ...appData.global, ...next }, appData.global);
    appData.inspectors.forEach((inspector) => {
      const previous = getInspectorSettings(inspector);
      inspector.settings = normalizeSettings({ ...previous, ...next }, previous);
    });
    saveData();
    if (shouldRender) render();
    return;
  }

  if (!target.matches("[data-main-field]")) return;
  const inspector = appData.inspectors[Number(target.dataset.index)];
  if (!inspector) return;
  const field = target.dataset.mainField;
  if (field === "name") inspector.name = target.value;
  if (field === "series") inspector.series = target.value;
  if (field === "schoolsCount") inspector.schoolsCount = numberOrZero(target.value);
  saveData();
  if (shouldRender) render();
}

function createInspector() {
  const settings = getSummarySettings();
  const id = appData.inspectors.reduce((max, inspector) => Math.max(max, numberOrZero(inspector.id)), 0) + 1;
  const index = appData.inspectors.length;
  const inspector = {
    id,
    name: `سەرپەرشتیاری نوێ ${id}`,
    series: String(100 + index),
    schoolsCount: 0,
    schoolsInCharge: { supervision: 0, externalEval: 0, total: 0 },
    schoolsPlanned: { supervision: 0, externalEval: 0, total: 0 },
    schoolsVisited: { supervision: 0, externalEval: 0, other: 0, total: 0 },
    stats: defaultStats(),
    settings: normalizeSettings(settings, appData.global),
    monthlyActivities: {},
    activities: [],
  };
  regenerateInspectorActivities(inspector);
  return inspector;
}

function addInspector() {
  const inspector = createInspector();
  appData.inspectors.push(inspector);
  saveData();
  currentView = "summary";
  render();
  showToast("سەرپەرشتیاری نوێ زیادکرا.");
}

function handleMainPageClick(event) {
  const removeButton = event.target.closest("[data-remove-inspector]");
  if (!removeButton) return;
  removeInspector(Number(removeButton.dataset.removeInspector));
}

function removeInspector(index) {
  if (appData.inspectors.length <= 1) {
    showToast("نابێت دوا سەرپەرشتیار بسڕدرێتەوە.");
    return;
  }
  const [removed] = appData.inspectors.splice(index, 1);
  if (!removed) return;
  if (currentInspectorId === removed.id) {
    currentInspectorId = appData.inspectors[0].id;
    currentView = "summary";
  }
  normalizeData(appData);
  saveData();
  render();
  showToast("سەرپەرشتیار سڕدرایەوە.");
}

function updateSchoolCount(event, shouldRender = true) {
  if (!event.target.matches("input")) return;
  const inspector = getCurrentInspector();
  if (!inspector) return;
  const group = event.target.dataset.group;
  const field = event.target.dataset.field;
  inspector[group][field] = numberOrZero(event.target.value);
  inspector.schoolsInCharge = normalizeTriple(inspector.schoolsInCharge);
  inspector.schoolsPlanned = normalizeTriple(inspector.schoolsPlanned);
  inspector.schoolsVisited = normalizeVisited(inspector.schoolsVisited);
  saveData();
  if (shouldRender) renderInspector();
}

function updateStatCount(event, shouldRender = true) {
  if (!event.target.matches("input")) return;
  const inspector = getCurrentInspector();
  if (!inspector) return;
  inspector.stats[event.target.dataset.field] = numberOrZero(event.target.value);
  saveData();
  if (shouldRender) renderSummary();
}

function updateActivity(event) {
  if (!event.target.matches("input")) return;
  const inspector = getCurrentInspector();
  const index = Number(event.target.dataset.index);
  inspector.activities[index].activity = event.target.value;
  inspector.activities[index].isHoliday = event.target.value.trim() === "پشوو";
  saveInspectorActivitiesForSettings(inspector);
  saveData();
}

function getCurrentInspector() {
  return appData.inspectors.find((inspector) => inspector.id === currentInspectorId);
}

function calculateTotals() {
  return appData.inspectors.reduce(
    (totals, inspector) => {
      const stats = inspector.stats;
      totals.schoolsCount += numberOrZero(inspector.schoolsCount);
      totals.inChargeSupervision += inspector.schoolsInCharge.supervision;
      totals.inChargeExternal += inspector.schoolsInCharge.externalEval;
      totals.inChargeTotal += inspector.schoolsInCharge.total;
      totals.plannedSupervision += inspector.schoolsPlanned.supervision;
      totals.plannedExternal += inspector.schoolsPlanned.externalEval;
      totals.plannedTotal += inspector.schoolsPlanned.total;
      totals.visitedSupervision += inspector.schoolsVisited.supervision;
      totals.visitedExternal += inspector.schoolsVisited.externalEval;
      totals.visitedOther += inspector.schoolsVisited.other;
      totals.visitedTotal += inspector.schoolsVisited.total;
      statFields.forEach(([field]) => {
        totals[field] += numberOrZero(stats[field]);
      });
      return totals;
    },
    {
      schoolsCount: 0,
      inChargeSupervision: 0,
      inChargeExternal: 0,
      inChargeTotal: 0,
      plannedSupervision: 0,
      plannedExternal: 0,
      plannedTotal: 0,
      visitedSupervision: 0,
      visitedExternal: 0,
      visitedOther: 0,
      visitedTotal: 0,
      ...Object.fromEntries(statFields.map(([field]) => [field, 0])),
    },
  );
}

function printAllSheets() {
  normalizeData(appData);
  saveData();
  const printRoot = document.getElementById("printRoot");
  printRoot.innerHTML = buildPrintDocument();
  showToast("خشتەکان بۆ چاپ ئامادەکران.");
  setTimeout(() => waitForPrintImages().then(() => window.print()), 80);
}

function printSummarySheet() {
  normalizeData(appData);
  saveData();
  const printRoot = document.getElementById("printRoot");
  printRoot.innerHTML = buildSummaryPrintPage();
  showToast("پوختە بۆ چاپ ئامادەکرا.");
  setTimeout(() => waitForPrintImages().then(() => window.print()), 80);
}

function printCurrentInspector() {
  const inspector = getCurrentInspector();
  if (!inspector) return;
  readInspectorSettingsFromFields(inspector);
  normalizeData(appData);
  saveData();
  const printRoot = document.getElementById("printRoot");
  const index = appData.inspectors.findIndex((item) => item.id === inspector.id);
  printRoot.innerHTML = buildDailyPrintPage(inspector, index);
  showToast("خشتەی ئەم سەرپەرشتیارە بۆ چاپ ئامادەکرا.");
  setTimeout(() => waitForPrintImages().then(() => window.print()), 80);
}

function buildPrintDocument() {
  const inspectorPages = appData.inspectors
    .map((inspector, index) => buildDailyPrintPage(inspector, index))
    .join("");
  return `${inspectorPages}${buildSummaryPrintPage()}`;
}

function buildDailyPrintPage(inspector, index) {
  const settings = getInspectorSettings(inspector);
  const signatureImage = getInspectorSignatureImage(inspector);
  const year = Number(settings.activityYear);
  const month = Number(settings.month);
  const layout = pickLayout(getDaysInMonth(year, month));
  const rows = [];
  for (let day = 1; day <= layout.capacity; day += 1) {
    const activity = inspector.activities[day - 1];
    rows.push(`<tr class="green">
      <td class="rtl">${activity ? escapeHtml(activity.day) : "&nbsp;"}</td>
      <td class="ltr">${activity ? escapeHtml(formatDisplayDate(activity.date)) : "&nbsp;"}</td>
      <td colspan="10" class="rtl activity-text">${activity ? escapeHtml(activity.activity) : "&nbsp;"}</td>
    </tr>`);
  }

  const sic = inspector.schoolsInCharge;
  const planned = inspector.schoolsPlanned;
  const visited = inspector.schoolsVisited;
  const stats = inspector.stats;

  return `<section class="print-page">
    <div class="daily-print-frame">
    <table class="excel-sheet excel-daily">
      <colgroup>
        <col style="width: 11.57%">
        <col style="width: 13.10%">
        <col style="width: 6.46%">
        <col style="width: 6.69%">
        <col style="width: 6.46%">
        <col style="width: 21.26%">
        <col style="width: 7.46%">
        <col style="width: 6.46%">
        <col style="width: 3.83%">
        <col style="width: 6.46%">
        <col style="width: 6.46%">
        <col style="width: 4.95%">
      </colgroup>
      <tbody>
        <tr class="top-title"><td colspan="12" class="rtl">بەڕێوەبەرایەتی دڵنیایی جۆری و سەرپەرشتیکردنی پەروەردەیی هەولێر</td></tr>
        <tr class="top-title"><td colspan="12" class="rtl">یەکەی دڵنیایی جۆریی وسەرپەرشتیکردنی پەروەردەیی هەولێر</td></tr>
        <tr class="top-title">
          <td class="rtl">مانگی </td>
          <td class="ltr meta-value">${escapeHtml(month)}</td>
          <td class="rtl">ساڵی </td>
          <td class="ltr meta-value study-year-value">${escapeHtml(settings.studyYear)}</td>
          <td colspan="3" class="rtl">خشتەی کارو چالاکی مانگانە</td>
          <td class="rtl">زنجیرە</td>
          <td colspan="2" class="ltr meta-value">${escapeHtml(inspector.series)}</td>
          <td colspan="2">&nbsp;</td>
        </tr>
        <tr class="green">
          <th class="rtl">ڕۆژ</th>
          <th class="rtl">ڕێکەوت</th>
          <th colspan="10" class="rtl">کورتەی کارو چالاکی ئەنجامدراو</th>
        </tr>
        ${rows.join("")}
        <tr class="green group-header">
          <td colspan="3" class="rtl">ژمارەی ئەو قوتابخانانەی لە ئەستۆیەتی</td>
          <td colspan="3" class="rtl">ژمارەی ئەو قوتابخانانەی لە پلاندانە بۆ سەردانیکردن لەم مانگە</td>
          <td colspan="6" class="rtl">ژمارەی ئەو قوتابخانانەی سەردانی کردووە لەم مانگە</td>
        </tr>
        <tr class="green sub-header">
          <td class="rtl">سەرپەرشتیاری پەروەردەیی</td>
          <td class="rtl">هەلسەنگاندنی دەرەکی</td>
          <td class="rtl">کۆ</td>
          <td class="rtl">سەرپەرشتیاری پەروەردەیی</td>
          <td class="rtl">هەلسەنگاندنی دەرەکی</td>
          <td class="rtl">کۆ</td>
          <td class="rtl">سەرپەرشتیاری پەروەردەیی</td>
          <td class="rtl">هەلسەنگاندنی دەرەکی</td>
          <td colspan="2" class="rtl">هەر سەردانێکی تری فەرمی</td>
          <td colspan="2" class="rtl">کۆ</td>
        </tr>
        <tr class="green value-row">
          <td>${sic.supervision}</td>
          <td>${sic.externalEval}</td>
          <td>${sic.total}</td>
          <td>${planned.supervision}</td>
          <td>${planned.externalEval}</td>
          <td>${planned.total}</td>
          <td>${visited.supervision}</td>
          <td>${visited.externalEval}</td>
          <td colspan="2">${visited.other || ""}</td>
          <td colspan="2">${visited.total}</td>
        </tr>
        <tr class="green stat-row stat-row-1">
          <td colspan="3" class="rtl">ژمارەی کۆبونەوەکانی بەشداریکردوە</td>
          <td>${optionalPrintNumber(stats.meetingsParticipated)}</td>
          <td colspan="4" class="rtl">ژمارەی کۆبونەوەکانی ئەنجامی داوە</td>
          <td>${optionalPrintNumber(stats.meetingsCompleted)}</td>
          <td colspan="2" class="rtl">ژمارەی لیژنەکانی بەشداریکردووە</td>
          <td>${optionalPrintNumber(stats.committees)}</td>
        </tr>
        <tr class="green stat-row stat-row-2">
          <td colspan="3" rowspan="2" class="rtl">ژمارەی لێکۆڵینەوەکانی کە بەشداریکردووە</td>
          <td rowspan="2">${optionalPrintNumber(stats.research)}</td>
          <td colspan="4" class="rtl">ژمارەی ئەو وانە مەشقی و ڕاهێنانی کە خۆی ئەنجامی داوە</td>
          <td>${optionalPrintNumber(stats.trainingConducted)}</td>
          <td colspan="2" class="rtl">ژمارەی ئەو کۆرو سمیناری کە خۆی ئەنجامی داوە</td>
          <td>${optionalPrintNumber(stats.seminarsConducted)}</td>
        </tr>
        <tr class="green stat-row stat-row-3">
          <td colspan="4" class="rtl">ژمارەی بەشداریکردنی لە کۆبونەوەی دایبابان</td>
          <td>${optionalPrintNumber(stats.parentMeetings)}</td>
          <td colspan="2" class="rtl">ژمارەی ئامادەبوونی لە وانەی مامۆستایی ڕاهێنەر</td>
          <td>${optionalPrintNumber(stats.teacherTrainingAttendance)}</td>
        </tr>
        <tr class="green stat-row stat-row-4">
          <td colspan="3" class="rtl">ژمارەی ڕاپۆرتەکانی لەسەر پرۆگرام و پرۆسەی پەروەردە</td>
          <td>${optionalPrintNumber(stats.reports)}</td>
          <td colspan="4" class="rtl">ژمارەی یاداشتەکانی</td>
          <td>${optionalPrintNumber(stats.notes)}</td>
          <td colspan="2" class="rtl">ژمارەی چالاکی جۆراو جۆری کە ئەنجامی داوە</td>
          <td>${optionalPrintNumber(stats.otherActivities)}</td>
        </tr>
        <tr class="stat-spacer">
          <td colspan="2">&nbsp;</td><td colspan="2">&nbsp;</td><td>&nbsp;</td><td colspan="2">&nbsp;</td>
          <td>&nbsp;</td><td colspan="2" class="signature-image-cell">${signatureImage}</td><td colspan="2">&nbsp;</td>
        </tr>
        <tr class="signature">
          <td colspan="2" class="rtl">ناوی سەرپەرشتیار:</td>
          <td colspan="2" class="rtl">${escapeHtml(inspector.name)}</td>
          <td class="rtl">بەروار</td>
          <td colspan="2" class="ltr">${escapeHtml(getLastDayDisplay(inspector))}</td>
          <td class="rtl">واژوو</td>
          <td colspan="2" class="stamp">&nbsp;</td>
          <td colspan="2">&nbsp;</td>
        </tr>
      </tbody>
    </table>
      <span class="daily-frame-border daily-frame-left" aria-hidden="true"></span>
      <span class="daily-frame-border daily-frame-right" aria-hidden="true"></span>
      <span class="daily-frame-border daily-frame-top" aria-hidden="true"></span>
      <span class="daily-frame-border daily-frame-bottom" aria-hidden="true"></span>
    </div>
  </section>`;
}

function buildSummaryPrintPage() {
  const totals = calculateTotals();
  const settings = getSummarySettings();
  const summaryColumns = [
    "3.0887%", "13.7768%", "11.0315%",
    "2.0584%", "2.9896%", "2.5493%", "2.9896%", "2.5493%", "2.9896%", "2.5493%",
    "2.9896%", "2.5493%", "2.9896%", "2.5493%", "2.5493%", "2.9896%", "2.9896%",
    "4.0683%", "2.4018%", "2.4018%", "2.4018%", "2.4018%", "2.9896%", "2.4018%",
    "2.4018%", "2.4018%", "2.9896%", "2.4018%", "2.1574%", "2.4023%",
  ];
  const inspectorRows = appData.inspectors
    .map((inspector, index) => buildSummaryPrintRow(index + 1, inspector))
    .join("");
  const totalRow = buildSummaryTotalRow(totals);
  const headSignature = getSummaryHeadSignature(settings.head);
  const headSignatureImage = headSignature
    ? `<img class="summary-head-signature-image" src="assets/signatures/${escapeHtml(headSignature.file)}?v=${SIGNATURE_ASSET_VERSION}" alt="">`
    : "";

  return `<section class="print-page summary-print-page">
    <table class="excel-sheet excel-summary" dir="rtl">
      <colgroup>${summaryColumns.map((width) => `<col style="width: ${width}">`).join("")}</colgroup>
      <tbody>
        <tr class="title-row"><td colspan="30" class="rtl sheet-title">یەكەی دڵنیایی جۆری سەرپەرشتیكردنی پەروەردەیی ناوەندی هەولێر</td></tr>
        <tr class="title-row">
          <td colspan="17" class="rtl">پوختەی ( كاری مانگانە / چاڵاكی مانگانە )ی سەرپەرشتیارانی پەروەردەیی / ناوەندی هەولێر مانگی</td>
          <td class="month-value ltr">${escapeHtml(settings.month)}</td>
          <td colspan="12" class="rtl year-cell">ساڵی ${escapeHtml(settings.studyYear)}</td>
        </tr>
        <tr class="header-row main-group-row">
          <td rowspan="5" class="rtl">ژ</td>
          <td rowspan="5" class="rtl">ناوى سەرپەرشتيار</td>
          <td rowspan="5" class="rtl vertical"><span>پسپۆرى</span></td>
          <td colspan="13" class="rtl">سەردانەكان</td>
          <td colspan="14" class="rtl">چالاكيەكان</td>
        </tr>
        <tr class="header-row subgroup-row">
          <td colspan="4" class="rtl">ژمارەى ئەو</td>
          <td colspan="4" class="rtl">ژمارەى ئەو قوتابخانانەى</td>
          <td colspan="5" class="rtl">ژمارەى سەردانەكانى</td>
          ${summaryVerticalHeaders()}
        </tr>
        <tr class="header-row subgroup-row">
          <td colspan="4" class="rtl">قوتابخانانەى لە</td>
          <td colspan="4" class="rtl">لە پلاندان بۆ سەردانيكردن</td>
          <td colspan="5" class="rtl">بۆ قوتابخانەكان</td>
        </tr>
        <tr class="header-row subgroup-row">
          <td colspan="4" class="rtl">ئەستۆيەتى</td>
          <td colspan="4" class="rtl">لەم مانگە</td>
          <td colspan="5" class="rtl">لەم مانگە</td>
        </tr>
        <tr class="header-row vertical-label-row">
          <td class="rtl vertical"><span>پسپۆرى</span></td>
          <td class="rtl vertical"><span>سەرپەرشتیاری پەروەردەیی</span></td>
          <td class="rtl vertical"><span>هەڵسەنگاندنى دەرەكى</span></td>
          <td class="rtl vertical"><span>كۆ</span></td>
          <td class="rtl vertical"><span>پسپۆرى</span></td>
          <td class="rtl vertical"><span>سەرپەرشتیاری پەروەردەیی</span></td>
          <td class="rtl vertical"><span>هەڵسەنگاندنى دەرەكى</span></td>
          <td class="rtl vertical"><span>كۆ</span></td>
          <td class="rtl vertical"><span>پسپۆرى</span></td>
          <td class="rtl vertical"><span>سەرپەرشتیاری پەروەردەیی</span></td>
          <td class="rtl vertical"><span>هەڵسەنگاندنى دەرەكى</span></td>
          <td class="rtl vertical"><span>هەر سەردانێكى ترى فەرمى</span></td>
          <td class="rtl vertical"><span>كۆ</span></td>
        </tr>
        ${inspectorRows}
        ${totalRow}
        <tr class="sign-spacer summary-signature-row"><td colspan="21">&nbsp;</td><td colspan="6" class="summary-head-signature-cell">${headSignatureImage}</td><td colspan="3">&nbsp;</td></tr>
        <tr class="sign-spacer"><td colspan="30">&nbsp;</td></tr>
        <tr class="sign-label-row"><td colspan="21" rowspan="2">&nbsp;</td><td colspan="6" rowspan="2" class="rtl">بەرپرسی بەش</td><td colspan="3" rowspan="2">&nbsp;</td></tr>
        <tr class="sign-label-continuation"></tr>
        <tr class="sign-name-row"><td colspan="21" rowspan="2">&nbsp;</td><td colspan="6" rowspan="2" class="rtl">${escapeHtml(settings.head)}</td><td colspan="3" rowspan="2">&nbsp;</td></tr>
        <tr class="sign-name-continuation"></tr>
      </tbody>
    </table>
  </section>`;
}

function summaryVerticalHeaders() {
  const headers = [
    "ژمارەی ئەو قوتابخانانەی لە ئەستۆیدایە",
    "ژمارەی کۆبونەوەکانى بەشداریکردووە",
    "ژمارەی لێکۆڵینەوەکانی کە بەشداریکردووە",
    "ژمارەی ڕاپۆرتەکانی لەسەر پرۆگرام و پڕۆسەی پەروەردە",
    "ژمارەی قوتابخانەکانی بە پێی پلانی ئەو مانگەی",
    "ژمارەی کۆبونەوەکانی ئەنجامی داوە",
    "ژمارەی ئەو وانە مەشقی و ڕاهێنانی کە خۆی ئەنجامی داوە",
    "ژمارەی بەشداریکردنی کۆبونەوەی دایبابان",
    "ژمارەی یاداشت",
    "ژمارەی سەرداناکان بەگشتی بۆ قوتابخانەکان",
    "ژمارەی لێژنەکانی بەشداریکردووە",
    "ژمارەی ئەو کۆڕو سیمیناری کە خۆی ئەنجامیداوە",
    "ژمارەی ئامادەبوونی لە وانەی مامۆستای ڕاهێنەر",
    "ژمارەی چاڵاکی جۆراوجۆری کە ئەنجامی داوە",
  ];
  return headers
    .map((header) => `<td rowspan="4" class="rtl vertical"><span>${escapeHtml(header)}</span></td>`)
    .join("");
}

function buildSummaryPrintRow(rowNumber, inspector) {
  const stats = inspector.stats;
  const settings = getInspectorSettings(inspector);
  const values = [
    rowNumber,
    inspector.name,
    settings.specialty,
    0,
    inspector.schoolsInCharge.supervision,
    inspector.schoolsInCharge.externalEval,
    inspector.schoolsInCharge.total,
    0,
    inspector.schoolsPlanned.supervision,
    inspector.schoolsPlanned.externalEval,
    inspector.schoolsPlanned.total,
    0,
    inspector.schoolsVisited.supervision,
    inspector.schoolsVisited.externalEval,
    inspector.schoolsVisited.other,
    inspector.schoolsVisited.total,
    inspector.schoolsCount,
    stats.meetingsParticipated,
    stats.research,
    stats.reports,
    inspector.schoolsPlanned.total,
    stats.meetingsCompleted,
    stats.trainingConducted,
    stats.parentMeetings,
    stats.notes,
    inspector.schoolsVisited.total,
    stats.committees,
    stats.seminarsConducted,
    stats.teacherTrainingAttendance,
    stats.otherActivities,
  ];
  return `<tr>${values.map((value, index) => printCell(value, index === 1 ? "name-cell" : "")).join("")}</tr>`;
}

function buildSummaryTotalRow(totals) {
  const values = [
    0,
    totals.inChargeSupervision,
    totals.inChargeExternal,
    totals.inChargeTotal,
    0,
    totals.plannedSupervision,
    totals.plannedExternal,
    totals.plannedTotal,
    0,
    totals.visitedSupervision,
    totals.visitedExternal,
    totals.visitedOther,
    totals.visitedTotal,
    totals.schoolsCount,
    totals.meetingsParticipated,
    totals.research,
    totals.reports,
    totals.plannedTotal,
    totals.meetingsCompleted,
    totals.trainingConducted,
    totals.parentMeetings,
    totals.notes,
    totals.visitedTotal,
    totals.committees,
    totals.seminarsConducted,
    totals.teacherTrainingAttendance,
    totals.otherActivities,
  ];
  return `<tr class="total"><td colspan="3" class="rtl">کۆى گشتى</td>${values
    .map((value) => printCell(value))
    .join("")}</tr>`;
}

function printCell(value, extraClass = "") {
  const isNumber = typeof value === "number";
  const className = [isNumber ? "ltr" : "rtl", extraClass].filter(Boolean).join(" ");
  return `<td class="${className}">${value === "" ? "&nbsp;" : escapeHtml(value)}</td>`;
}

function optionalPrintNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) && number !== 0 ? escapeHtml(number) : "";
}

function normalizeInspectorSignatureName(value) {
  return String(value || "")
    .normalize("NFC")
    .toLowerCase()
    .replace(/[\u064B-\u065F\u0670]/g, "")
    .replace(/[أإآٱ]/g, "ا")
    .replace(/[ىی]/g, "ي")
    .replace(/ک/g, "ك")
    .replace(/ـ/g, "")
    .replace(/\s+/g, "");
}

function getInspectorSignature(inspector) {
  const normalizedName = normalizeInspectorSignatureName(inspector?.name);
  return INSPECTOR_SIGNATURES.find((signature) =>
    signature.aliases.some((alias) => normalizedName.includes(normalizeInspectorSignatureName(alias))),
  ) || null;
}

function getSummaryHeadSignature(headName) {
  const normalizedHeadName = normalizeInspectorSignatureName(headName);
  if (!normalizedHeadName) return null;
  return INSPECTOR_SIGNATURES.find((signature) =>
    signature.aliases.some((alias) => {
      const normalizedAlias = normalizeInspectorSignatureName(alias);
      return normalizedAlias.includes(normalizedHeadName) || normalizedHeadName.includes(normalizedAlias);
    }),
  ) || null;
}

function getInspectorSignatureImage(inspector) {
  const signature = getInspectorSignature(inspector);
  if (!signature) return "";
  const src = `assets/signatures/${signature.file}?v=${SIGNATURE_ASSET_VERSION}`;
  return `<img class="inspector-signature-image" src="${escapeHtml(src)}" alt="">`;
}

async function generateWorkbook(options = {}) {
  if (!window.JSZip) {
    showToast("JSZip بارنەبووە. پەیوەندی ئینتەرنێت پێویستە.");
    return;
  }
  normalizeData(appData);
  if (options.persist !== false) saveData();
  showToast("Excel دروست دەکرێت...");
  try {
    const summarySettings = getSummarySettings();
    const year = Number(summarySettings.activityYear);
    const month = Number(summarySettings.month);
    const dailyLayouts = appData.inspectors.map((inspector) => {
      const settings = getInspectorSettings(inspector);
      return pickLayout(getDaysInMonth(Number(settings.activityYear), Number(settings.month)));
    });
    const template = await fetch(TEMPLATE_PATH).then((response) => {
      if (!response.ok) throw new Error("Template not found");
      return response.arrayBuffer();
    });
    const zip = await JSZip.loadAsync(template);
    const dailySheetInfos = await ensureDailyWorkbookSheets(zip, appData.inspectors.length);
    await removeWorkbookFillColors(zip);
    await normalizeDailySheetLayouts(zip, dailyLayouts, dailySheetInfos);
    await fillInputSheet(zip);
    await fillDailySheets(zip, dailyLayouts, dailySheetInfos);
    await embedInspectorSignatures(zip, dailyLayouts, dailySheetInfos);
    await fillSummarySheet(zip);
    await embedSummaryHeadSignature(zip);
    await updateWorkbookPrintAreas(zip, dailyLayouts, dailySheetInfos);
    await removeCalculationChain(zip);
    const blob = await zip.generateAsync({
      type: "blob",
      compression: "DEFLATE",
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    downloadBlob(blob, `inspection-monthly-${year}-${String(month).padStart(2, "0")}.xlsx`);
    showToast("فایلی Excel ئامادەیە.");
  } catch (error) {
    console.error(error);
    showToast("نەتوانرا فایلی Excel دروست بکرێت.");
  }
}

async function embedInspectorSignatures(zip, dailyLayouts, dailySheetInfos) {
  const contentTypesContext = await loadXmlContext(zip, "[Content_Types].xml");
  const contentTypesRoot = contentTypesContext.doc.documentElement;
  const drawingParts = new Set(
    Array.from(contentTypesRoot.children)
      .filter((node) => node.localName === "Override")
      .map((node) => node.getAttribute("PartName")),
  );

  for (let index = 0; index < appData.inspectors.length; index += 1) {
    const inspector = appData.inspectors[index];
    const signature = getInspectorSignature(inspector);
    if (!signature) continue;

    const imagePath = `assets/signatures/${signature.file}?v=${SIGNATURE_ASSET_VERSION}`;
    const imageResponse = await fetch(imagePath);
    if (!imageResponse.ok) throw new Error(`Signature image not found: ${signature.file}`);

    const mediaFile = `xl/media/inspector-signature-${index + 1}.png`;
    const drawingName = `signature-daily-${index + 1}.xml`;
    const drawingFile = `xl/drawings/${drawingName}`;
    const drawingRelsFile = `xl/drawings/_rels/${drawingName}.rels`;
    zip.file(mediaFile, await imageResponse.arrayBuffer());

    const drawingOverridePath = `/${drawingFile}`;
    if (!drawingParts.has(drawingOverridePath)) {
      const override = contentTypesContext.doc.createElementNS(CONTENT_TYPES_NS, "Override");
      override.setAttribute("PartName", drawingOverridePath);
      override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml");
      contentTypesRoot.appendChild(override);
      drawingParts.add(drawingOverridePath);
    }

    const info = dailySheetInfos[index];
    const worksheetContext = await loadXmlContext(zip, info.file);
    const worksheetRelsFile = getWorksheetRelsFile(info.file);
    const worksheetRelsContext = await loadXmlContext(zip, worksheetRelsFile);
    const worksheetRelsRoot = worksheetRelsContext.doc.documentElement;
    const drawingRelId = `rId${getNextRelationshipNumber(worksheetRelsRoot)}`;
    const drawingRel = worksheetRelsContext.doc.createElementNS(PACKAGE_REL_NS, "Relationship");
    drawingRel.setAttribute("Id", drawingRelId);
    drawingRel.setAttribute("Type", DRAWING_REL_TYPE);
    drawingRel.setAttribute("Target", `../drawings/${drawingName}`);
    worksheetRelsRoot.appendChild(drawingRel);

    const drawingElement = worksheetContext.doc.createElementNS(MAIN_NS, "drawing");
    if (!worksheetContext.doc.documentElement.getAttribute("xmlns:r")) {
      worksheetContext.doc.documentElement.setAttribute("xmlns:r", OFFICE_REL_NS);
    }
    drawingElement.setAttributeNS(OFFICE_REL_NS, "r:id", drawingRelId);
    const trailingSheetElements = new Set([
      "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst",
    ]);
    const insertBefore = Array.from(worksheetContext.doc.documentElement.children)
      .find((node) => trailingSheetElements.has(node.localName));
    worksheetContext.doc.documentElement.insertBefore(drawingElement, insertBefore || null);

    const imageRel = worksheetRelsContext.doc.implementation.createDocument(PACKAGE_REL_NS, "Relationships", null);
    const imageRelRoot = imageRel.documentElement;
    const imageRelationship = imageRel.createElementNS(PACKAGE_REL_NS, "Relationship");
    imageRelationship.setAttribute("Id", "rId1");
    imageRelationship.setAttribute("Type", IMAGE_REL_TYPE);
    imageRelationship.setAttribute("Target", `../media/${mediaFile.split("/").pop()}`);
    imageRelRoot.appendChild(imageRelationship);
    zip.file(drawingRelsFile, serializeXml(imageRel));
    zip.file(drawingFile, buildInspectorSignatureDrawingXml(signature, dailyLayouts[index], "rId1"));
    zip.file(info.file, serializeXml(worksheetContext.doc));
    zip.file(worksheetRelsFile, serializeXml(worksheetRelsContext.doc));
  }

  zip.file(contentTypesContext.file, serializeXml(contentTypesContext.doc));
}

function buildInspectorSignatureDrawingXml(signature, layout, imageRelId, options = {}) {
  const maxWidthEmu = 1_600_000;
  const maxHeightEmu = 500_000;
  const aspect = signature.width / signature.height;
  const cx = Math.min(maxWidthEmu, Math.round(maxHeightEmu * aspect));
  const cy = Math.round(cx / aspect);
  const spacerRow = Math.max(0, Number(layout?.signatureRow || 2) - 2);
  const anchorCol = Math.max(0, Number(options.col ?? 8));
  const colOffset = Math.max(0, Number(options.colOffset ?? 0));
  const rowOffset = Math.max(0, Number(options.rowOffset ?? 30000));
  const imageName = `Inspector signature ${signature.file.replace(/\.png$/i, "")}`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="${DRAWING_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${OFFICE_REL_NS}"><xdr:oneCellAnchor><xdr:from><xdr:col>${anchorCol}</xdr:col><xdr:colOff>${colOffset}</xdr:colOff><xdr:row>${spacerRow}</xdr:row><xdr:rowOff>${rowOffset}</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="${imageName}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${imageRelId}" cstate="print"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`;
}

async function embedSummaryHeadSignature(zip) {
  const signature = getSummaryHeadSignature(getSummarySettings().head);
  if (!signature) return;

  const imageResponse = await fetch(`assets/signatures/${signature.file}?v=${SIGNATURE_ASSET_VERSION}`);
  if (!imageResponse.ok) throw new Error(`Head signature image not found: ${signature.file}`);
  zip.file("xl/media/summary-head-signature.png", await imageResponse.arrayBuffer());

  const contentTypesContext = await loadXmlContext(zip, "[Content_Types].xml");
  const contentTypesRoot = contentTypesContext.doc.documentElement;
  const drawingFile = "xl/drawings/summary-head-signature.xml";
  const drawingOverridePath = `/${drawingFile}`;
  const hasDrawingOverride = Array.from(contentTypesRoot.children).some((node) =>
    node.localName === "Override" && node.getAttribute("PartName") === drawingOverridePath,
  );
  if (!hasDrawingOverride) {
    const override = contentTypesContext.doc.createElementNS(CONTENT_TYPES_NS, "Override");
    override.setAttribute("PartName", drawingOverridePath);
    override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml");
    contentTypesRoot.appendChild(override);
  }

  const worksheetContext = await loadXmlContext(zip, SUMMARY_SHEET_FILE);
  const worksheetRelsFile = getWorksheetRelsFile(SUMMARY_SHEET_FILE);
  const worksheetRelsContext = await loadXmlContext(zip, worksheetRelsFile);
  const worksheetRelsRoot = worksheetRelsContext.doc.documentElement;
  const drawingRelId = `rId${getNextRelationshipNumber(worksheetRelsRoot)}`;
  const drawingRel = worksheetRelsContext.doc.createElementNS(PACKAGE_REL_NS, "Relationship");
  drawingRel.setAttribute("Id", drawingRelId);
  drawingRel.setAttribute("Type", DRAWING_REL_TYPE);
  drawingRel.setAttribute("Target", "../drawings/summary-head-signature.xml");
  worksheetRelsRoot.appendChild(drawingRel);

  const drawingElement = worksheetContext.doc.createElementNS(MAIN_NS, "drawing");
  if (!worksheetContext.doc.documentElement.getAttribute("xmlns:r")) {
    worksheetContext.doc.documentElement.setAttribute("xmlns:r", OFFICE_REL_NS);
  }
  drawingElement.setAttributeNS(OFFICE_REL_NS, "r:id", drawingRelId);
  const trailingSheetElements = new Set([
    "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst",
  ]);
  const insertBefore = Array.from(worksheetContext.doc.documentElement.children)
    .find((node) => trailingSheetElements.has(node.localName));
  worksheetContext.doc.documentElement.insertBefore(drawingElement, insertBefore || null);

  const imageRel = worksheetRelsContext.doc.implementation.createDocument(PACKAGE_REL_NS, "Relationships", null);
  const imageRelRoot = imageRel.documentElement;
  const imageRelationship = imageRel.createElementNS(PACKAGE_REL_NS, "Relationship");
  imageRelationship.setAttribute("Id", "rId1");
  imageRelationship.setAttribute("Type", IMAGE_REL_TYPE);
  imageRelationship.setAttribute("Target", "../media/summary-head-signature.png");
  imageRelRoot.appendChild(imageRelationship);

  const signRow = Math.max(17, 8 + appData.inspectors.length + 5);
  const drawingXml = buildInspectorSignatureDrawingXml(signature, { signatureRow: signRow - 3 }, "rId1", { col: 21, colOffset: 570000, rowOffset: 0 });
  zip.file(drawingFile, drawingXml);
  zip.file("xl/drawings/_rels/summary-head-signature.xml.rels", serializeXml(imageRel));
  zip.file(worksheetContext.file, serializeXml(worksheetContext.doc));
  zip.file(worksheetRelsFile, serializeXml(worksheetRelsContext.doc));
  zip.file(contentTypesContext.file, serializeXml(contentTypesContext.doc));
}

function pickLayout(days) {
  if (days >= 31) {
    return {
      capacity: 31,
      baseFile: "xl/worksheets/sheet3.xml",
      countRow: 38,
      statRow: 39,
      signatureRow: 44,
      printEndRow: 47,
    };
  }
  if (days >= 29) {
    return {
      capacity: 30,
      baseFile: "xl/worksheets/sheet4.xml",
      countRow: 37,
      statRow: 38,
      signatureRow: 43,
      printEndRow: 48,
    };
  }
  return {
    capacity: 28,
    baseFile: "xl/worksheets/sheet2.xml",
    countRow: 35,
    statRow: 36,
    signatureRow: 41,
    printEndRow: 43,
  };
}

function getDailySheetInfos(count) {
  const dailyCount = Math.max(1, count);
  return Array.from({ length: dailyCount }, (_, index) => ({
    file: getDailySheetFile(index),
    name: getDailySheetName(index),
    index,
  }));
}

function getDailySheetFile(index) {
  return dailySheetFiles[index] || `xl/worksheets/sheet${index + 3}.xml`;
}

function getDailySheetName(index) {
  return dailySheetNames[index] || `س${toArabicIndicDigits(index + 1)}`;
}

function getWorksheetRelsFile(sheetFile) {
  const fileName = sheetFile.split("/").pop();
  return `xl/worksheets/_rels/${fileName}.rels`;
}

function toArabicIndicDigits(value) {
  const digits = "٠١٢٣٤٥٦٧٨٩";
  return String(value).replace(/\d/g, (digit) => digits[Number(digit)]);
}

async function ensureDailyWorkbookSheets(zip, inspectorCount) {
  const infos = getDailySheetInfos(inspectorCount);
  if (inspectorCount <= dailySheetFiles.length) return infos;

  await ensureWorkbookSheetEntries(zip, infos.slice(dailySheetFiles.length));
  await ensureWorksheetContentTypes(zip, infos.slice(dailySheetFiles.length));
  await updateDocPropsSheetNames(zip, infos);
  return infos;
}

async function ensureWorkbookSheetEntries(zip, newInfos) {
  const workbookContext = await loadXmlContext(zip, "xl/workbook.xml");
  const workbookDoc = workbookContext.doc;
  const sheets = firstChildByName(workbookDoc.documentElement, "sheets");
  const summarySheet = Array.from(sheets.children).find((sheet) => sheet.getAttribute("name") === SUMMARY_SHEET_NAME);
  const existingNames = new Set(Array.from(sheets.children).map((sheet) => sheet.getAttribute("name")));
  const maxSheetId = Array.from(sheets.children).reduce(
    (max, sheet) => Math.max(max, numberOrZero(sheet.getAttribute("sheetId"))),
    0,
  );

  const relsContext = await loadXmlContext(zip, "xl/_rels/workbook.xml.rels");
  const relsRoot = relsContext.doc.documentElement;
  const existingTargets = new Set(
    Array.from(relsRoot.children)
      .filter((rel) => rel.getAttribute("Type") === WORKSHEET_REL_TYPE)
      .map((rel) => rel.getAttribute("Target")),
  );
  let nextSheetId = maxSheetId + 1;
  let nextRelNumber = getNextRelationshipNumber(relsRoot);

  newInfos.forEach((info) => {
    const target = info.file.replace(/^xl\//, "");
    let relId = findRelationshipIdByTarget(relsRoot, target);
    if (!relId) {
      relId = `rId${nextRelNumber}`;
      nextRelNumber += 1;
      const rel = relsContext.doc.createElementNS(PACKAGE_REL_NS, "Relationship");
      rel.setAttribute("Id", relId);
      rel.setAttribute("Type", WORKSHEET_REL_TYPE);
      rel.setAttribute("Target", target);
      relsRoot.appendChild(rel);
      existingTargets.add(target);
    }

    if (!existingNames.has(info.name)) {
      const sheet = workbookDoc.createElementNS(MAIN_NS, "sheet");
      sheet.setAttribute("name", info.name);
      sheet.setAttribute("sheetId", String(nextSheetId));
      sheet.setAttributeNS(OFFICE_REL_NS, "r:id", relId);
      nextSheetId += 1;
      sheets.insertBefore(sheet, summarySheet || null);
      existingNames.add(info.name);
    }
  });

  zip.file(workbookContext.file, serializeXml(workbookDoc));
  zip.file(relsContext.file, serializeXml(relsContext.doc));
}

async function ensureWorksheetContentTypes(zip, newInfos) {
  const context = await loadXmlContext(zip, "[Content_Types].xml");
  const root = context.doc.documentElement;
  const existingParts = new Set(
    Array.from(root.children)
      .filter((node) => node.localName === "Override")
      .map((node) => node.getAttribute("PartName")),
  );
  newInfos.forEach((info) => {
    const partName = `/${info.file}`;
    if (existingParts.has(partName)) return;
    const override = context.doc.createElementNS(CONTENT_TYPES_NS, "Override");
    override.setAttribute("PartName", partName);
    override.setAttribute("ContentType", WORKSHEET_CONTENT_TYPE);
    root.appendChild(override);
    existingParts.add(partName);
  });
  zip.file(context.file, serializeXml(context.doc));
}

async function updateDocPropsSheetNames(zip, dailyInfos) {
  const file = "docProps/app.xml";
  if (!zip.file(file)) return;
  const context = await loadXmlContext(zip, file);
  const root = context.doc.documentElement;
  const titles = firstChildByName(root, "TitlesOfParts");
  if (!titles) return;
  const vector = firstChildByName(titles, "vector");
  if (!vector) return;
  Array.from(vector.childNodes).forEach((node) => node.parentNode.removeChild(node));
  const names = ["سەرەتا", ...dailyInfos.map((info) => info.name), SUMMARY_SHEET_NAME];
  names.forEach((name) => {
    const item = context.doc.createElementNS(vector.namespaceURI, "vt:lpstr");
    item.textContent = name;
    vector.appendChild(item);
  });
  vector.setAttribute("size", String(names.length));

  const headingPairs = firstChildByName(root, "HeadingPairs");
  const headingVector = headingPairs ? firstChildByName(headingPairs, "vector") : null;
  if (headingVector) {
    const variants = Array.from(headingVector.children).filter((node) => node.localName === "variant");
    const worksheetCount = variants[1] ? firstChildByName(variants[1], "i4") || firstChildByName(variants[1], "int") : null;
    if (worksheetCount) worksheetCount.textContent = String(names.length);
  }
  zip.file(context.file, serializeXml(context.doc));
}

function getNextRelationshipNumber(relsRoot) {
  return (
    Array.from(relsRoot.children).reduce((max, rel) => {
      const match = /^rId(\d+)$/.exec(rel.getAttribute("Id") || "");
      return match ? Math.max(max, Number(match[1])) : max;
    }, 0) + 1
  );
}

function findRelationshipIdByTarget(relsRoot, target) {
  const rel = Array.from(relsRoot.children).find(
    (node) => node.getAttribute("Type") === WORKSHEET_REL_TYPE && node.getAttribute("Target") === target,
  );
  return rel?.getAttribute("Id") || "";
}

async function normalizeDailySheetLayouts(zip, layouts, dailySheetInfos = getDailySheetInfos(appData.inspectors.length)) {
  for (let index = 0; index < dailySheetInfos.length; index += 1) {
    const layout = layouts[index] || layouts[0] || pickLayout(31);
    const baseXml = await zip.file(layout.baseFile).async("string");
    const info = dailySheetInfos[index];
    zip.file(info.file, baseXml);
    const baseRels = getWorksheetRelsFile(layout.baseFile);
    const targetRels = getWorksheetRelsFile(info.file);
    if (zip.file(baseRels)) {
      const relsXml = await zip.file(baseRels).async("string");
      zip.file(targetRels, relsXml);
    } else if (targetRels !== baseRels && zip.file(targetRels)) {
      zip.remove(targetRels);
    }
  }
}

async function removeWorkbookFillColors(zip) {
  const context = await loadXmlContext(zip, "xl/styles.xml");
  const doc = context.doc;
  const styleSheet = doc.documentElement;
  let fills = firstChildByName(styleSheet, "fills");
  if (!fills) {
    fills = doc.createElementNS(MAIN_NS, "fills");
    styleSheet.insertBefore(fills, firstChildByName(styleSheet, "borders") || null);
  }
  Array.from(fills.childNodes).forEach((node) => node.parentNode.removeChild(node));
  fills.appendChild(createFill(doc, "none"));
  fills.appendChild(createFill(doc, "gray125"));
  fills.setAttribute("count", "2");

  ["cellStyleXfs", "cellXfs"].forEach((listName) => {
    const list = firstChildByName(styleSheet, listName);
    if (!list) return;
    Array.from(list.children)
      .filter((node) => node.localName === "xf")
      .forEach((xf) => {
        xf.setAttribute("fillId", "0");
        xf.removeAttribute("applyFill");
      });
  });

  const dxfs = firstChildByName(styleSheet, "dxfs");
  if (dxfs) {
    Array.from(dxfs.getElementsByTagNameNS(MAIN_NS, "fill")).forEach((fill) => fill.parentNode.removeChild(fill));
  }
  zip.file(context.file, serializeXml(context.doc));
}

function createFill(doc, patternType) {
  const fill = doc.createElementNS(MAIN_NS, "fill");
  const patternFill = doc.createElementNS(MAIN_NS, "patternFill");
  patternFill.setAttribute("patternType", patternType);
  fill.appendChild(patternFill);
  return fill;
}

async function removeCalculationChain(zip) {
  zip.remove("xl/calcChain.xml");

  const contentTypesFile = "[Content_Types].xml";
  if (zip.file(contentTypesFile)) {
    const context = await loadXmlContext(zip, contentTypesFile);
    Array.from(context.doc.documentElement.children)
      .filter((node) => node.localName === "Override" && node.getAttribute("PartName") === "/xl/calcChain.xml")
      .forEach((node) => node.parentNode.removeChild(node));
    zip.file(context.file, serializeXml(context.doc));
  }

  const workbookRelsFile = "xl/_rels/workbook.xml.rels";
  if (zip.file(workbookRelsFile)) {
    const context = await loadXmlContext(zip, workbookRelsFile);
    Array.from(context.doc.documentElement.children)
      .filter((node) => node.localName === "Relationship" && /calcChain$/i.test(node.getAttribute("Type") || ""))
      .forEach((node) => node.parentNode.removeChild(node));
    zip.file(context.file, serializeXml(context.doc));
  }
}

async function fillInputSheet(zip) {
  const context = await loadSheetContext(zip, "xl/worksheets/sheet1.xml");
  const settings = getSummarySettings();
  prepareInputInspectorRows(context, appData.inspectors.length);
  setText(context, "B2", settings.specialty);
  setText(context, "C2", settings.studyYear);
  setNumber(context, "D2", settings.month);
  setText(context, "F2", settings.head);
  appData.inspectors.forEach((inspector, index) => {
    const row = 5 + index;
    setNumber(context, `A${row}`, index + 1);
    setText(context, `B${row}`, inspector.name);
    setNumber(context, `C${row}`, numberOrZero(inspector.series));
    setNumber(context, `D${row}`, inspector.schoolsCount);
  });
  clearInputInspectorRows(context, 5 + appData.inspectors.length, 8);
  ensureSheetDimension(context.doc, "J", 4 + appData.inspectors.length);
  saveSheetContext(zip, context);
}

function prepareInputInspectorRows(context, inspectorCount) {
  const lastRow = 4 + inspectorCount;
  for (let row = 9; row <= lastRow; row += 1) {
    copyRowFormat(context, 8, row);
  }
}

function clearInputInspectorRows(context, startRow, endRow) {
  for (let row = startRow; row <= endRow; row += 1) {
    ["A", "B", "C", "D"].forEach((column) => clearCell(context, `${column}${row}`));
  }
}

async function fillDailySheets(zip, layouts, dailySheetInfos = getDailySheetInfos(appData.inspectors.length)) {
  const dailyCount = Math.min(dailySheetInfos.length, appData.inspectors.length);
  for (let index = 0; index < dailyCount; index += 1) {
    const context = await loadSheetContext(zip, dailySheetInfos[index].file);
    fillDailySheet(context, appData.inspectors[index], layouts[index] || layouts[0]);
    saveSheetContext(zip, context);
  }
}

function fillDailySheet(context, inspector, layout) {
  const settings = getInspectorSettings(inspector);
  const year = Number(settings.activityYear);
  const month = Number(settings.month);
  setNumber(context, "B3", month);
  setText(context, "D3", settings.studyYear);
  setNumber(context, "I3", numberOrZero(inspector.series));
  for (let day = 1; day <= layout.capacity; day += 1) {
    const row = 4 + day;
    const activity = inspector.activities[day - 1];
    if (activity) {
      setText(context, `A${row}`, activity.day);
      setNumber(context, `B${row}`, excelDateSerial(year, month, day));
      setText(context, `C${row}`, activity.activity);
    } else {
      clearCell(context, `A${row}`);
      clearCell(context, `B${row}`);
      clearCell(context, `C${row}`);
    }
  }
  const countRow = layout.countRow;
  const sic = inspector.schoolsInCharge;
  const sp = inspector.schoolsPlanned;
  const sv = inspector.schoolsVisited;
  setNumber(context, `A${countRow}`, sic.supervision);
  setNumber(context, `B${countRow}`, sic.externalEval);
  setNumber(context, `C${countRow}`, sic.total);
  setNumber(context, `D${countRow}`, sp.supervision);
  setNumber(context, `E${countRow}`, sp.externalEval);
  setNumber(context, `F${countRow}`, sp.total);
  setNumber(context, `G${countRow}`, sv.supervision);
  setNumber(context, `H${countRow}`, sv.externalEval);
  setOptionalNumber(context, `I${countRow}`, sv.other);
  setNumber(context, `K${countRow}`, sv.total);

  const statRow = layout.statRow;
  const stats = inspector.stats;
  setOptionalNumber(context, `D${statRow}`, stats.meetingsParticipated);
  setOptionalNumber(context, `I${statRow}`, stats.meetingsCompleted);
  setOptionalNumber(context, `L${statRow}`, stats.committees);
  setOptionalNumber(context, `D${statRow + 1}`, stats.research);
  setOptionalNumber(context, `I${statRow + 1}`, stats.trainingConducted);
  setOptionalNumber(context, `L${statRow + 1}`, stats.seminarsConducted);
  setOptionalNumber(context, `I${statRow + 2}`, stats.parentMeetings);
  setOptionalNumber(context, `L${statRow + 2}`, stats.teacherTrainingAttendance);
  setOptionalNumber(context, `D${statRow + 3}`, stats.reports);
  setOptionalNumber(context, `I${statRow + 3}`, stats.notes);
  setOptionalNumber(context, `L${statRow + 3}`, stats.otherActivities);
  setText(context, `B${layout.signatureRow}`, inspector.name);
  setText(context, `F${layout.signatureRow}`, getLastDayDisplay(inspector));
}

async function fillSummarySheet(zip) {
  const context = await loadSheetContext(zip, SUMMARY_SHEET_FILE);
  const totals = calculateTotals();
  const settings = getSummarySettings();
  prepareSummaryRows(context, appData.inspectors.length);
  setText(context, "R2", String(settings.month));
  setText(context, "S2", `ساڵی ${settings.studyYear}`);
  appData.inspectors.forEach((inspector, index) => {
    const row = 8 + index;
    writeSummaryRow(context, row, index + 1, inspector);
  });
  const totalRow = 8 + appData.inspectors.length;
  setText(context, `A${totalRow}`, "کۆى گشتى");
  const totalValues = {
    D: 0,
    E: totals.inChargeSupervision,
    F: totals.inChargeExternal,
    G: totals.inChargeTotal,
    H: 0,
    I: totals.plannedSupervision,
    J: totals.plannedExternal,
    K: totals.plannedTotal,
    L: 0,
    M: totals.visitedSupervision,
    N: totals.visitedExternal,
    O: totals.visitedOther,
    P: totals.visitedTotal,
    Q: totals.schoolsCount,
    R: totals.meetingsParticipated,
    S: totals.research,
    T: totals.reports,
    U: totals.plannedTotal,
    V: totals.meetingsCompleted,
    W: totals.trainingConducted,
    X: totals.parentMeetings,
    Y: totals.notes,
    Z: totals.visitedTotal,
    AA: totals.committees,
    AB: totals.seminarsConducted,
    AC: totals.teacherTrainingAttendance,
    AD: totals.otherActivities,
  };
  Object.entries(totalValues).forEach(([column, value]) => setNumber(context, `${column}${totalRow}`, value));
  const signRow = Math.max(17, totalRow + 5);
  if (signRow !== 17) clearCell(context, "V17");
  setText(context, `V${signRow}`, settings.head);
  clearSummaryRows(context, totalRow + 1, 12);
  ensureSheetDimension(context.doc, "AD", signRow);
  saveSheetContext(zip, context);
}

function prepareSummaryRows(context, inspectorCount) {
  const totalRow = 8 + inspectorCount;
  if (totalRow !== 12) {
    copyRowFormat(context, 12, totalRow);
  }
  for (let row = 12; row < totalRow; row += 1) {
    copyRowFormat(context, 11, row);
  }
  setMergeRange(context.doc, `A${totalRow}:C${totalRow}`, /^A\d+:C\d+$/);
}

function copyRowFormat(context, sourceRowNumber, targetRowNumber) {
  const sourceRow = context.rows.get(sourceRowNumber);
  if (!sourceRow) return;
  const targetRow = getOrCreateRow(context, targetRowNumber);
  Array.from(targetRow.attributes).forEach((attribute) => {
    if (attribute.name !== "r") targetRow.removeAttribute(attribute.name);
  });
  Array.from(sourceRow.attributes).forEach((attribute) => {
    if (attribute.name === "r") return;
    if (attribute.namespaceURI) targetRow.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
    else targetRow.setAttribute(attribute.name, attribute.value);
  });
  targetRow.setAttribute("r", String(targetRowNumber));

  Array.from(targetRow.getElementsByTagNameNS(MAIN_NS, "c")).forEach((targetCell) => {
    context.cells.delete(targetCell.getAttribute("r"));
    targetCell.parentNode.removeChild(targetCell);
  });
  Array.from(sourceRow.getElementsByTagNameNS(MAIN_NS, "c")).forEach((sourceCell) => {
    const column = sourceCell.getAttribute("r").match(/[A-Z]+/)?.[0];
    if (!column) return;
    const targetCell = context.doc.createElementNS(MAIN_NS, "c");
    targetCell.setAttribute("r", `${column}${targetRowNumber}`);
    Array.from(sourceCell.attributes).forEach((attribute) => {
      if (attribute.name === "r") return;
      if (attribute.namespaceURI) targetCell.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
      else targetCell.setAttribute(attribute.name, attribute.value);
    });
    clearCellNode(targetCell);
    targetRow.appendChild(targetCell);
    context.cells.set(targetCell.getAttribute("r"), targetCell);
  });
}

function clearSummaryRows(context, startRow, endRow) {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let column = 1; column <= getColumnNumber("AD"); column += 1) {
      clearCell(context, `${getColumnLetters(column)}${row}`);
    }
  }
}

function setMergeRange(doc, mergeRef, removePattern) {
  let mergeCells = firstChildByName(doc.documentElement, "mergeCells");
  if (!mergeCells) {
    mergeCells = doc.createElementNS(MAIN_NS, "mergeCells");
    const sheetData = firstChildByName(doc.documentElement, "sheetData");
    doc.documentElement.insertBefore(mergeCells, sheetData?.nextSibling || null);
  }
  Array.from(mergeCells.children)
    .filter((node) => node.localName === "mergeCell" && removePattern.test(node.getAttribute("ref") || ""))
    .forEach((node) => node.parentNode.removeChild(node));
  const merge = doc.createElementNS(MAIN_NS, "mergeCell");
  merge.setAttribute("ref", mergeRef);
  mergeCells.appendChild(merge);
  mergeCells.setAttribute("count", String(Array.from(mergeCells.children).length));
}

function writeSummaryRow(context, row, index, inspector) {
  const stats = inspector.stats;
  const settings = getInspectorSettings(inspector);
  const values = {
    A: index,
    B: inspector.name,
    C: settings.specialty,
    D: 0,
    E: inspector.schoolsInCharge.supervision,
    F: inspector.schoolsInCharge.externalEval,
    G: inspector.schoolsInCharge.total,
    H: 0,
    I: inspector.schoolsPlanned.supervision,
    J: inspector.schoolsPlanned.externalEval,
    K: inspector.schoolsPlanned.total,
    L: 0,
    M: inspector.schoolsVisited.supervision,
    N: inspector.schoolsVisited.externalEval,
    O: inspector.schoolsVisited.other,
    P: inspector.schoolsVisited.total,
    Q: inspector.schoolsCount,
    R: stats.meetingsParticipated,
    S: stats.research,
    T: stats.reports,
    U: inspector.schoolsPlanned.total,
    V: stats.meetingsCompleted,
    W: stats.trainingConducted,
    X: stats.parentMeetings,
    Y: stats.notes,
    Z: inspector.schoolsVisited.total,
    AA: stats.committees,
    AB: stats.seminarsConducted,
    AC: stats.teacherTrainingAttendance,
    AD: stats.otherActivities,
  };
  Object.entries(values).forEach(([column, value]) => {
    if (typeof value === "string") setText(context, `${column}${row}`, value);
    else setNumber(context, `${column}${row}`, value);
  });
}

function getSheetLocalIds(workbookDoc) {
  const sheets = firstChildByName(workbookDoc.documentElement, "sheets");
  const localIds = new Map();
  Array.from(sheets?.children || []).forEach((sheet, index) => {
    localIds.set(sheet.getAttribute("name"), index);
  });
  return localIds;
}

async function updateWorkbookPrintAreas(zip, layouts, dailySheetInfos = getDailySheetInfos(appData.inspectors.length)) {
  const context = await loadXmlContext(zip, "xl/workbook.xml");
  const doc = context.doc;
  const sheetLocalIds = getSheetLocalIds(doc);
  let definedNames = firstChildByName(doc.documentElement, "definedNames");
  if (!definedNames) {
    definedNames = doc.createElementNS(MAIN_NS, "definedNames");
    doc.documentElement.appendChild(definedNames);
  }
  Array.from(definedNames.children)
    .filter((node) => node.getAttribute("name") === "_xlnm.Print_Area")
    .forEach((node) => node.parentNode.removeChild(node));
  dailySheetInfos.forEach((info, index) => {
    const layout = layouts[index] || layouts[0] || pickLayout(31);
    const node = doc.createElementNS(MAIN_NS, "definedName");
    node.setAttribute("name", "_xlnm.Print_Area");
    node.setAttribute("localSheetId", String(sheetLocalIds.get(info.name) ?? index + 1));
    node.textContent = `'${info.name}'!$A$1:$L$${layout.printEndRow}`;
    definedNames.appendChild(node);
  });
  const summaryArea = doc.createElementNS(MAIN_NS, "definedName");
  summaryArea.setAttribute("name", "_xlnm.Print_Area");
  summaryArea.setAttribute("localSheetId", String(sheetLocalIds.get(SUMMARY_SHEET_NAME) ?? dailySheetInfos.length + 1));
  summaryArea.textContent = `'${SUMMARY_SHEET_NAME}'!$A$1:$AD$${Math.max(18, appData.inspectors.length + 13)}`;
  definedNames.appendChild(summaryArea);
  let calcPr = firstChildByName(doc.documentElement, "calcPr");
  if (!calcPr) {
    calcPr = doc.createElementNS(MAIN_NS, "calcPr");
    doc.documentElement.appendChild(calcPr);
  }
  calcPr.setAttribute("calcMode", "auto");
  calcPr.setAttribute("fullCalcOnLoad", "1");
  calcPr.setAttribute("forceFullCalc", "1");
  zip.file(context.file, serializeXml(doc));
}

async function loadSheetContext(zip, file) {
  const context = await loadXmlContext(zip, file);
  context.rows = new Map();
  context.cells = new Map();
  const rows = Array.from(context.doc.getElementsByTagNameNS(MAIN_NS, "row"));
  rows.forEach((row) => {
    context.rows.set(Number(row.getAttribute("r")), row);
    Array.from(row.getElementsByTagNameNS(MAIN_NS, "c")).forEach((cell) => {
      context.cells.set(cell.getAttribute("r"), cell);
    });
  });
  return context;
}

async function loadXmlContext(zip, file) {
  const xml = await zip.file(file).async("string");
  return { file, doc: parseXml(xml) };
}

function saveSheetContext(zip, context) {
  zip.file(context.file, serializeXml(context.doc));
}

function parseXml(xml) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) {
    throw new Error("Invalid workbook XML");
  }
  return doc;
}

function serializeXml(doc) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${new XMLSerializer().serializeToString(doc.documentElement)}`;
}

function getOrCreateCell(context, ref, styleRef = null) {
  if (context.cells.has(ref)) return context.cells.get(ref);
  const rowNumber = getRowNumber(ref);
  const columnNumber = getColumnNumber(ref);
  const row = getOrCreateRow(context, rowNumber);
  const cell = context.doc.createElementNS(MAIN_NS, "c");
  cell.setAttribute("r", ref);
  const styleSource = styleRef ? context.cells.get(styleRef) : findNearestStyleCell(context, rowNumber, columnNumber);
  if (styleSource?.getAttribute("s")) cell.setAttribute("s", styleSource.getAttribute("s"));
  const cells = Array.from(row.getElementsByTagNameNS(MAIN_NS, "c"));
  const next = cells.find((item) => getColumnNumber(item.getAttribute("r")) > columnNumber);
  row.insertBefore(cell, next || null);
  context.cells.set(ref, cell);
  return cell;
}

function getOrCreateRow(context, rowNumber) {
  if (context.rows.has(rowNumber)) return context.rows.get(rowNumber);
  const sheetData = context.doc.getElementsByTagNameNS(MAIN_NS, "sheetData")[0];
  const row = context.doc.createElementNS(MAIN_NS, "row");
  row.setAttribute("r", String(rowNumber));
  const rows = Array.from(sheetData.getElementsByTagNameNS(MAIN_NS, "row"));
  const next = rows.find((item) => Number(item.getAttribute("r")) > rowNumber);
  sheetData.insertBefore(row, next || null);
  context.rows.set(rowNumber, row);
  return row;
}

function findNearestStyleCell(context, rowNumber, columnNumber) {
  const row = context.rows.get(rowNumber);
  if (!row) return null;
  const cells = Array.from(row.getElementsByTagNameNS(MAIN_NS, "c"));
  let nearest = null;
  let distance = Infinity;
  cells.forEach((cell) => {
    const diff = Math.abs(getColumnNumber(cell.getAttribute("r")) - columnNumber);
    if (cell.getAttribute("s") && diff < distance) {
      nearest = cell;
      distance = diff;
    }
  });
  return nearest;
}

function setText(context, ref, value, styleRef = null) {
  const cell = getOrCreateCell(context, ref, styleRef);
  clearCellNode(cell);
  const text = value == null ? "" : String(value);
  if (!text) return;
  cell.setAttribute("t", "inlineStr");
  const inlineString = context.doc.createElementNS(MAIN_NS, "is");
  const t = context.doc.createElementNS(MAIN_NS, "t");
  if (/^\s|\s$/.test(text)) t.setAttributeNS(XML_NS, "xml:space", "preserve");
  t.textContent = text;
  inlineString.appendChild(t);
  cell.appendChild(inlineString);
}

function setNumber(context, ref, value, styleRef = null) {
  const cell = getOrCreateCell(context, ref, styleRef);
  clearCellNode(cell);
  const number = Number(value);
  if (!Number.isFinite(number)) return;
  const v = context.doc.createElementNS(MAIN_NS, "v");
  v.textContent = String(number);
  cell.appendChild(v);
}

function setOptionalNumber(context, ref, value, styleRef = null) {
  const number = Number(value);
  if (!Number.isFinite(number) || number === 0) {
    clearCell(context, ref);
    return;
  }
  setNumber(context, ref, number, styleRef);
}

function clearCell(context, ref) {
  const cell = getOrCreateCell(context, ref);
  clearCellNode(cell);
}

function clearCellNode(cell) {
  Array.from(cell.childNodes).forEach((node) => node.parentNode.removeChild(node));
  cell.removeAttribute("t");
}

function firstChildByName(parent, localName) {
  return Array.from(parent.childNodes).find((node) => node.localName === localName) || null;
}

function getRowNumber(ref) {
  return Number(ref.match(/\d+/)?.[0] || 1);
}

function getColumnNumber(ref) {
  const letters = ref.match(/[A-Z]+/)?.[0] || "A";
  return letters.split("").reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0);
}

function getColumnLetters(number) {
  let value = number;
  let letters = "";
  while (value > 0) {
    value -= 1;
    letters = String.fromCharCode(65 + (value % 26)) + letters;
    value = Math.floor(value / 26);
  }
  return letters || "A";
}

function ensureSheetDimension(doc, endColumn, endRow) {
  let dimension = firstChildByName(doc.documentElement, "dimension");
  if (!dimension) {
    dimension = doc.createElementNS(MAIN_NS, "dimension");
    doc.documentElement.insertBefore(dimension, doc.documentElement.firstChild || null);
  }
  const currentRef = dimension.getAttribute("ref") || "A1:A1";
  const currentEnd = currentRef.split(":").pop() || currentRef;
  const currentEndColumn = currentEnd.match(/[A-Z]+/)?.[0] || "A";
  const currentEndRow = Number(currentEnd.match(/\d+/)?.[0] || 1);
  const nextEndColumn = getColumnLetters(Math.max(getColumnNumber(currentEndColumn), getColumnNumber(endColumn)));
  const nextEndRow = Math.max(currentEndRow, endRow);
  dimension.setAttribute("ref", `A1:${nextEndColumn}${nextEndRow}`);
}

function excelDateSerial(year, month, day) {
  const excelEpoch = Date.UTC(1899, 11, 30);
  const current = Date.UTC(year, month - 1, day);
  return Math.round((current - excelEpoch) / 86400000);
}

function downloadJson() {
  const settings = getSummarySettings();
  const blob = new Blob([JSON.stringify(appData, null, 2)], { type: "application/json" });
  downloadBlob(blob, `inspection-data-${settings.activityYear}-${String(settings.month).padStart(2, "0")}.json`);
}

function importJson(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    const parsed = safeJsonParse(String(reader.result));
    if (!parsed) {
      showToast("فایلی JSON دروست نییە.");
      return;
    }
    appData = parsed;
    normalizeData(appData);
    saveData();
    currentView = "summary";
    render();
    showToast("زانیارییەکان هێنرانەوە.");
  };
  reader.readAsText(file);
  event.target.value = "";
}

function resetData() {
  if (!confirm("دڵنیایت لە پاککردنەوەی هەموو زانیارییەکان؟")) return;
  appData = clone(defaultData);
  normalizeData(appData);
  saveData();
  currentView = "summary";
  render();
  showToast("زانیارییەکان نوێکرانەوە.");
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function formatDisplayDate(isoDate) {
  const [year, month, day] = isoDate.split("-");
  return `${day}/${month}/${year}`;
}

function getLastDayDisplay(inspector = null) {
  const settings = inspector ? getInspectorSettings(inspector) : getSummarySettings();
  const year = Number(settings.activityYear) || new Date().getFullYear();
  const month = Number(settings.month) || 1;
  const day = getDaysInMonth(year, month);
  return `${String(day).padStart(2, "0")}/${String(month).padStart(2, "0")}/${year}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

let toastTimer = null;
function showToast(message) {
  const toast = document.getElementById("toast");
  toast.textContent = message;
  toast.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove("show"), 2600);
}

init();
