const state = {
  workbook: null,
  sheetName: null,
  rawRows: [],
  tasks: [],
  filtered: [],
  ganttTaskLookup: new Map(),
  context: { buckets: new Map(), users: new Map(), goals: new Map(), planName: "" },
  ganttCols: { task: 312, owner: 142, schedule: 168 },
  filtersCollapsed: false,
  lang: "zh"
};

/* ─── i18n ─── */
const i18n = {
  zh: {
    heroHeading: "把 Planner Excel 變成可執行的專案作戰台",
    heroCopy: "自動識別 Planner 匯出的 Tasks、Consolidated Data、Buckets、Users 工作表，修復 ID 顯示問題；按專案分組、負責人和狀態組織任務，未設置 Due Date 的任務仍會進入欄目與清單，但不會繪製甘特條。",
    badge1: "本機解析", badge2: "可摺疊甘特層級", badge3: "PMO 風險視圖",
    uploadBtn: "選擇 Planner Excel",
    uploadHint: "支援 Planner 預設多工作表匯出",
    noticeDefault: "請上傳 Microsoft Planner 匯出的 Excel。建議保留 Tasks、Buckets、Users 和 Consolidated Data 工作表。",
    commandLabel: "Command", filterTitle: "篩選與視圖", resetBtn: "重置",
    filterCollapse: "收起",
    filterExpand: "展開",
    sectionFileTime: "檔案與時間",
    labelSheet: "任務資料來源", sheetPlaceholder: "等待上傳",
    sheetHint: "優先讀取 Consolidated Data；選擇 Tasks 時會自動用 Buckets / Users 映射名稱。",
    labelWindow: "視窗範圍",
    labelGroup1: "甘特一級分組", labelGroup2: "甘特二級分組",
    labelSearch: "全域搜尋", searchPlaceholder: "任務、負責人、分組、標籤、備註...",
    optBucket: "專案分組 / Bucket", optAssigned: "負責人 / Owner",
    optStatus: "狀態 / Status", optPriority: "優先級 / Priority",
    optDueWeek: "截止週 / Due Week", optNone: "不分組",
    sectionBucket: "專案分組", sectionAssignee: "負責人", sectionStatus: "狀態",
    exportBtn: "匯出當前 CSV",
    statTotal: "任務總數", statTotalNote: "當前篩選範圍",
    statCompletion: "完成率",
    statOverdue: "逾期風險", statOverdueNote: "未完成且超過截止日",
    statNoDue: "未排期", statNoDueNote: "無 Due Date，進入欄目但不畫條",
    tabGantt: "甘特圖", tabAnalytics: "任務分析", tabPmo: "PMO 總覽", tabTable: "任務清單",
    ganttTitle: "可摺疊層級甘特圖",
    ganttDesc: "一級分組可展開/收起；二級按負責人或分組組織。無 Due Date 的任務保留在所在層級，用「未排期」行顯示，不繪製時間條。",
    focusMode: "專注模式", focusModeExit: "退出專注",
    densityBtn: "緊湊視圖", densityBtnAlt: "舒適視圖",
    expandAll: "展開全部", collapseAll: "收起全部",
    legendUrgent: "Urgent 緊急", legendImportant: "Important 重要",
    legendMedium: "Medium 中等", legendLow: "Low 低",
    ganttMeta: "上傳 Excel 後顯示專案排期。",
    analyticsTitle: "任務分析",
    analyticsDesc: "按優先級、截止時間和負責人維度分析任務分佈，輔助專案決策。",
    analyticsPriority: "優先級分佈", analyticsAssignee: "負責人完成率", analyticsWeekly: "每週到期任務",
    pmoTitle: "PMO 總覽",
    pmoDesc: "聚焦能指導專案經理行動的指標：組合健康、分組進度、負責人負荷和風險清單。",
    pmoHealth: "組合健康", pmoWorkload: "負責人負荷", pmoBucket: "分組進度", pmoRisk: "需要關注",
    thTask: "任務名稱", thBucket: "專案分組", thAssignee: "負責人",
    thStatus: "狀態", thPriority: "優先級",
    thStart: "開始", thDue: "截止", thCompleted: "完成", thLabels: "標籤 / Checklist",
    taskCount: "{n} 條任務", completedNote: "{n} 個已完成",
    noData: "暫無資料", noRisk: "當前篩選範圍沒有明顯排期風險。",
    emptyGantt: "上傳 Planner Excel 後生成可摺疊甘特圖",
    emptyFilter: "沒有符合篩選條件的任務。",
    libLoading: "Excel 解析庫尚未載入完成，請重新整理頁面或檢查網路後重試。",
    parsing: "正在解析 Excel，並建立 Buckets / Users 映射…",
    noSheets: "Excel 中沒有可讀取的工作表。",
    noTaskSheets: "沒有找到包含 Task Name / 任務名稱 的 Planner 任務工作表。",
    sheetNoData: "這個工作表沒有可用任務資料。請確認選擇了 Planner 匯出的 Tasks 或 Consolidated Data 工作表。",
    loadedMsg: "已讀取 {n} 條任務{plan}。{mapping}。所有內容只在瀏覽器本機處理。",
    mappedBuckets: "{n} 條 Bucket ID 已映射", mappedUsers: "{n} 條使用者 ID 已映射",
    noMapping: "未發現需要映射的 ID",
    overdue: "逾期", dueSoon: "即將到期", noDueDate: "無 Due Date",
    unscheduled: "未排期", unassigned: "未分配", ungrouped: "未分組",
    unknown: "未知", allTasks: "全部任務",
    completionLabel: "完成率",
    healthCopy: "當前視圖共 {total} 條任務，{active} 條未完成。優先處理逾期、7 天內到期和無截止日任務。",
    healthDone: "完成", healthActive: "執行中", health7d: "7天內到期",
    healthOverdue: "逾期", healthNoDue: "未排期",
    riskItems: "條 · 風險", riskDone: "完成",
    ganttBaseYear: "基準年", ganttTasks: "條任務", ganttScheduled: "條已排期",
    ganttCoverage: "覆蓋率", ganttSections: "個 Section",
    ganttResizeHint: "拖動左側表頭分隔線可像 Excel 一樣調整列寬",
    noDueGanttNote: "個無 Due Date：保留在左側任務表，不畫時間條",
    noStartNote: "個無開始日期：按截止日定位",
    todayLabel: "今天",
    statusCompleted: "已完成", statusOverdue: "逾期", statusDueSoon: "即將到期",
    statusInProgress: "進行中", statusNotStarted: "未開始", statusUnknown: "未知",
    langToggle: "EN",
    unnamed: "未命名任務"
  },
  en: {
    heroHeading: "Turn Planner Excel into an Executable Project Warboard",
    heroCopy: "Auto-detect Planner-exported Tasks, Consolidated Data, Buckets, Users sheets; fix ID display issues; group by project, owner, and status. Tasks without Due Date still appear in lists but won't render Gantt bars.",
    badge1: "Local Parsing", badge2: "Collapsible Gantt", badge3: "PMO Risk View",
    uploadBtn: "Choose Planner Excel",
    uploadHint: "Supports Planner default multi-sheet export",
    noticeDefault: "Please upload a Microsoft Planner exported Excel. Keep Tasks, Buckets, Users, and Consolidated Data sheets.",
    commandLabel: "Command", filterTitle: "Filters & Views", resetBtn: "Reset",
    filterCollapse: "Hide",
    filterExpand: "Show",
    sectionFileTime: "File & Time",
    labelSheet: "Task Data Source", sheetPlaceholder: "Awaiting upload",
    sheetHint: "Prefers Consolidated Data; choosing Tasks auto-maps via Buckets / Users.",
    labelWindow: "Window Range",
    labelGroup1: "Gantt Primary Group", labelGroup2: "Gantt Secondary Group",
    labelSearch: "Global Search", searchPlaceholder: "Task, owner, group, label, notes...",
    optBucket: "Project Group / Bucket", optAssigned: "Owner / Assignee",
    optStatus: "Status", optPriority: "Priority",
    optDueWeek: "Due Week", optNone: "No Grouping",
    sectionBucket: "Project Groups", sectionAssignee: "Assignees", sectionStatus: "Status",
    exportBtn: "Export Current CSV",
    statTotal: "Total Tasks", statTotalNote: "Current filter scope",
    statCompletion: "Completion",
    statOverdue: "Overdue Risk", statOverdueNote: "Incomplete & past due",
    statNoDue: "Unscheduled", statNoDueNote: "No Due Date, listed but no bar",
    tabGantt: "Gantt Chart", tabAnalytics: "Analytics", tabPmo: "PMO Overview", tabTable: "Task List",
    ganttTitle: "Collapsible Hierarchical Gantt",
    ganttDesc: "Primary groups expand/collapse; secondary by owner or group. Tasks without Due Date remain in their tier shown as 'Unscheduled', no time bar drawn.",
    focusMode: "Focus Mode", focusModeExit: "Exit Focus",
    densityBtn: "Compact View", densityBtnAlt: "Comfortable View",
    expandAll: "Expand All", collapseAll: "Collapse All",
    legendUrgent: "Urgent", legendImportant: "Important",
    legendMedium: "Medium", legendLow: "Low",
    ganttMeta: "Upload Excel to view project schedule.",
    analyticsTitle: "Task Analytics",
    analyticsDesc: "Analyze task distribution by priority, due date, and assignee to support project decisions.",
    analyticsPriority: "Priority Distribution", analyticsAssignee: "Assignee Completion",
    analyticsWeekly: "Weekly Due Tasks",
    pmoTitle: "PMO Overview",
    pmoDesc: "Focus on actionable PM metrics: portfolio health, group progress, assignee workload, and risk list.",
    pmoHealth: "Portfolio Health", pmoWorkload: "Assignee Workload",
    pmoBucket: "Group Progress", pmoRisk: "Needs Attention",
    thTask: "Task Name", thBucket: "Project Group", thAssignee: "Assignee",
    thStatus: "Status", thPriority: "Priority",
    thStart: "Start", thDue: "Due", thCompleted: "Completed", thLabels: "Labels / Checklist",
    taskCount: "{n} tasks", completedNote: "{n} completed",
    noData: "No data", noRisk: "No significant scheduling risks in current filter.",
    emptyGantt: "Upload Planner Excel to generate collapsible Gantt",
    emptyFilter: "No tasks match current filters.",
    libLoading: "Excel parsing library not loaded. Please refresh or check network.",
    parsing: "Parsing Excel and building Buckets / Users mapping…",
    noSheets: "No readable sheets found in the Excel file.",
    noTaskSheets: "No Planner task sheet found with Task Name column.",
    sheetNoData: "This sheet has no usable task data. Ensure you selected a Planner Tasks or Consolidated Data sheet.",
    loadedMsg: "Loaded {n} tasks{plan}. {mapping}. All processing done locally in browser.",
    mappedBuckets: "{n} Bucket IDs mapped", mappedUsers: "{n} User IDs mapped",
    noMapping: "No ID mapping needed",
    overdue: "Overdue", dueSoon: "Due Soon", noDueDate: "No Due Date",
    unscheduled: "Unscheduled", unassigned: "Unassigned", ungrouped: "Ungrouped",
    unknown: "Unknown", allTasks: "All Tasks",
    completionLabel: "Completion",
    healthCopy: "Current view: {total} tasks, {active} incomplete. Prioritize overdue, due within 7 days, and unscheduled tasks.",
    healthDone: "Done", healthActive: "Active", health7d: "Due in 7d",
    healthOverdue: "Overdue", healthNoDue: "No Due",
    riskItems: "items · risk", riskDone: "done",
    ganttBaseYear: "Base Year", ganttTasks: "tasks", ganttScheduled: "scheduled",
    ganttCoverage: "coverage", ganttSections: "sections",
    ganttResizeHint: "Drag header dividers to resize columns like Excel",
    noDueGanttNote: "no Due Date: listed but no bar",
    noStartNote: "no start date: positioned by due date",
    todayLabel: "Today",
    statusCompleted: "Completed", statusOverdue: "Overdue", statusDueSoon: "Due Soon",
    statusInProgress: "In Progress", statusNotStarted: "Not Started", statusUnknown: "Unknown",
    langToggle: "中文",
    unnamed: "Unnamed Task"
  }
};

function t(key) { return (i18n[state.lang] || i18n.zh)[key] || key; }

const palette = {
  Urgent: "#D92D20",
  Important: "#F79009",
  Medium: "#0054A6",
  Normal: "#0054A6",
  Low: "#8EA4BA",
  "": "#0054A6"
};

const els = {};

const taskColumnCandidates = [
  "Task Name", "Task name", "任务名称", "Title", "标题", "Name", "名称"
];

const taskSheetPreference = ["Consolidated Data", "Tasks", "Task", "任务"];
const PROJECT_YEAR = 2027;

document.addEventListener("DOMContentLoaded", () => {
  cacheElements();
  bindEvents();
  clearVisuals();
  applyLangToUI();
  setNotice(t("noticeDefault"));
});

function cacheElements() {
  [
    "fileInput", "fileLabel", "sheetSelect", "bucketFilter", "assigneeFilter",
    "statusFilter", "startDate", "endDate", "subGroupBy", "groupBy", "resetBtn", "csvBtn",
    "notice", "totalTasks", "completionRate", "completedNote", "overdueTasks",
    "noDueTasks", "searchInput", "bucketCount", "assigneeCount", "statusCount",
    "ganttChart", "healthPanel", "workloadPanel", "bucketPanel", "riskPanel",
    "expandAllBtn", "collapseAllBtn", "focusModeBtn", "densityBtn", "tableCount", "langBtn",
    "filterCollapseBtn"
  ].forEach(id => { els[id] = document.getElementById(id); });
}

function bindEvents() {
  els.fileInput.addEventListener("change", handleFile);
  els.sheetSelect.addEventListener("change", () => loadSheet(els.sheetSelect.value));
  ["bucketFilter", "assigneeFilter", "statusFilter"]
    .forEach(id => els[id].addEventListener("change", applyFiltersAndRender));
  ["startDate", "endDate", "groupBy", "subGroupBy", "searchInput"]
    .forEach(id => els[id].addEventListener("input", applyFiltersAndRender));
  els.resetBtn.addEventListener("click", resetFilters);
  els.csvBtn.addEventListener("click", exportCsv);
  els.expandAllBtn.addEventListener("click", () => setGanttOpen(true));
  els.collapseAllBtn.addEventListener("click", () => setGanttOpen(false));
  els.focusModeBtn.addEventListener("click", toggleFocusMode);
  els.densityBtn.addEventListener("click", toggleGanttDensity);
  if (els.langBtn) els.langBtn.addEventListener("click", toggleLang);
  if (els.filterCollapseBtn) els.filterCollapseBtn.addEventListener("click", toggleFilterPanel);

  document.querySelectorAll(".tab-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".panel").forEach(panel => panel.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(`tab-${btn.dataset.tab}`).classList.add("active");
    });
  });
}

function toggleLang() {
  state.lang = state.lang === "zh" ? "en" : "zh";
  applyLangToUI();
  if (state.filtered.length) renderAll();
  else clearVisuals();
}

function applyLangToUI() {
  // Update all data-i18n elements
  document.querySelectorAll("[data-i18n]").forEach(el => {
    el.textContent = t(el.dataset.i18n);
  });
  document.querySelectorAll("[data-i18n-placeholder]").forEach(el => {
    el.placeholder = t(el.dataset.i18nPlaceholder);
  });
  // Lang toggle button text
  if (els.langBtn) els.langBtn.textContent = t("langToggle");
  // Select options
  document.querySelectorAll("[data-i18n-option]").forEach(opt => {
    opt.textContent = t(opt.dataset.i18nOption);
  });
  // Update notice if no data loaded
  if (!state.tasks.length) {
    setNotice(t("noticeDefault"));
  }
  syncFilterPanelToggle();
}

async function handleFile(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  if (!window.XLSX) {
    setNotice(t("libLoading"), "error");
    return;
  }

  try {
    setNotice(t("parsing"));
    const data = await file.arrayBuffer();
    state.workbook = XLSX.read(data, { type: "array", cellDates: true });
    state.context = buildPlannerContext(state.workbook);
    els.fileLabel.textContent = file.name;

    const sheets = state.workbook.SheetNames || [];
    if (!sheets.length) throw new Error(t("noSheets"));

    const taskSheets = findTaskSheets(state.workbook);
    if (!taskSheets.length) throw new Error(t("noTaskSheets"));

    const selected = chooseDefaultTaskSheet(taskSheets);
    fillSelect(els.sheetSelect, taskSheets, selected);
    els.sheetSelect.disabled = false;
    loadSheet(els.sheetSelect.value);
  } catch (error) {
    console.error(error);
    setNotice(`${error.message}`, "error");
    clearVisuals();
  }
}

function findTaskSheets(workbook) {
  return workbook.SheetNames.filter(name => {
    const rows = sheetRows(workbook, name);
    return rows.length && Object.keys(rows[0] || {}).some(key =>
      taskColumnCandidates.some(candidate => normalizeHeader(key) === normalizeHeader(candidate))
    );
  });
}

function chooseDefaultTaskSheet(taskSheets) {
  for (const preferred of taskSheetPreference) {
    const found = taskSheets.find(name => normalizeHeader(name) === normalizeHeader(preferred));
    if (found) return found;
  }
  return taskSheets[0];
}

function loadSheet(sheetName) {
  if (!state.workbook || !sheetName) return;
  state.sheetName = sheetName;
  state.rawRows = sheetRows(state.workbook, sheetName);
  state.tasks = normalizePlannerRows(state.rawRows, state.context);

  if (!state.tasks.length) {
    setNotice(t("sheetNoData"), "error");
    clearVisuals();
    return;
  }

  initializeFilters();
  applyFiltersAndRender();

  const mappedBuckets = state.tasks.filter(t => t.bucketWasMapped).length;
  const mappedUsers = state.tasks.filter(t => t.userWasMapped).length;
  const mappingNote = [
    mappedBuckets ? t("mappedBuckets").replace("{n}", mappedBuckets) : "",
    mappedUsers ? t("mappedUsers").replace("{n}", mappedUsers) : ""
  ].filter(Boolean).join("；");
  const plan = state.context.planName ? ` · ${state.context.planName}` : "";
  setNotice(t("loadedMsg").replace("{n}", state.tasks.length).replace("{plan}", plan).replace("{mapping}", mappingNote || t("noMapping")), "ok");
}

function sheetRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  return sheet ? XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false }) : [];
}

function buildPlannerContext(workbook) {
  const buckets = new Map();
  const users = new Map();
  const goals = new Map();
  let planName = "";

  const planSheet = findSheet(workbook, ["Plan", "计划"]);
  if (planSheet) {
    const row = sheetRows(workbook, planSheet)[0] || {};
    planName = cleanText(pick(row, ["Plan name", "Plan Name", "计划名称", "Name", "名称"]));
  }

  const bucketSheet = findSheet(workbook, ["Buckets", "Bucket", "存储桶", "分组"]);
  sheetRows(workbook, bucketSheet).forEach(row => {
    const id = cleanText(pick(row, ["Bucket ID", "Bucket Id", "BucketID", "ID", "存储桶 ID", "分组 ID"]));
    const name = readableGroup(pick(row, ["Bucket Name", "Bucket name", "Bucket", "存储桶名称", "分组名称", "Name", "名称"]));
    if (id && name) buckets.set(id, name);
  });

  const userSheet = findSheet(workbook, ["Users", "User", "用户"]);
  sheetRows(workbook, userSheet).forEach(row => {
    const id = cleanText(pick(row, ["User ID", "User Id", "UserID", "ID", "用户 ID"]));
    const name = cleanText(pick(row, ["User Name", "User name", "User", "用户名称", "用户名", "Name", "名称"]));
    const email = cleanText(pick(row, ["Email", "Mail", "邮箱", "电子邮件"]));
    const display = name ? titleCaseLoose(name) : prettifyUser(email);
    if (id && display) users.set(id, display);
    if (email && display) users.set(email.toLowerCase(), display);
  });

  const goalSheet = findSheet(workbook, ["Goals", "Goal", "目标"]);
  sheetRows(workbook, goalSheet).forEach(row => {
    const id = cleanText(pick(row, ["Goal ID", "Goal Id", "ID", "目标 ID"]));
    const name = cleanText(pick(row, ["Goal name", "Goal Name", "Goal", "目标名称", "Name", "名称"]));
    if (id && name) goals.set(id, name);
  });

  return { buckets, users, goals, planName };
}

function findSheet(workbook, candidates) {
  if (!workbook?.SheetNames) return "";
  for (const candidate of candidates) {
    const found = workbook.SheetNames.find(name => normalizeHeader(name) === normalizeHeader(candidate));
    if (found) return found;
  }
  return "";
}

function normalizePlannerRows(rows, context) {
  return rows.map((row, index) => {
    const createdDate = parseDate(pick(row, ["Created Date", "Created date", "Created", "创建日期", "建立日期"]));
    const startRaw = parseDate(pick(row, ["Start Date", "Start date", "Start", "开始日期", "开始时间"]));
    const dueDate = parseDate(pick(row, ["Due Date", "Due date", "Due", "截止日期", "到期日期", "到期时间"]));
    const completedDate = parseDate(pick(row, ["Completed Date", "Completed date", "Completed", "完成日期", "完成时间"]));
    const ganttStart = startRaw || dueDate || completedDate || createdDate;
    const ganttEnd = dueDate || completedDate || ganttStart;

    const taskName = readableProjectName(
      pick(row, ["Project Name", "项目名称", "Plan Name", "计划名称", "Task Name", "Task name", "任务名称", "Title", "标题", "Name", "名称"]),
      index
    );

    const bucketResult = resolveBucket(
      pick(row, ["Bucket Name", "Bucket name", "Bucket", "存储桶名称", "存储桶", "分组", "项目分组"]),
      context
    );
    const userResult = resolveUsers(
      pick(row, ["User Name", "User", "用户名称", "用户名", "负责人", "Assigned To", "Assigned to", "Assignees", "Owner", "分配给", "指派给", "被分配给"]),
      context
    );
    const createdByResult = resolveUsers(pick(row, ["Created By", "Created by", "创建者", "创建人"]), context);
    const completedByResult = resolveUsers(pick(row, ["Completed By", "Completed by", "完成人"]), context);

    const progress = cleanText(pick(row, ["Progress", "Status", "进度", "状态"])) || "Unknown";
    const priority = cleanText(pick(row, ["Priority", "Importance", "优先级", "重要性"])) || "Normal";
    const status = inferStatus(progress, dueDate, completedDate);
    const viewStartDate = toProjectYear(startRaw);
    const viewDueDate = toProjectYear(dueDate);
    const viewCompletedDate = toProjectYear(completedDate);
    const viewGanttStart = toProjectYear(ganttStart);
    const viewGanttEnd = toProjectYear(ganttEnd);
    const lateRaw = cleanText(pick(row, ["Late", "Is Late", "逾期"]));
    const checklistMeta = parseChecklistMeta(row);
    const labels = cleanText(pick(row, ["Labels", "Label names", "标签"]));
    const notes = cleanText(pick(row, ["Notes", "Description", "说明", "备注"]));
    const assignees = distinctPreserveOrder(userResult.value ? splitPersonNames(userResult.value) : []);

    return {
      rowIndex: index + 1,
      task: taskName,
      bucket: bucketResult.value || t("ungrouped"),
      assignedTo: assignees.join("; ") || t("unassigned"),
      assignees,
      progress,
      status,
      priority,
      createdBy: createdByResult.value || "Unknown",
      completedBy: completedByResult.value || "",
      createdDate,
      startDate: startRaw,
      dueDate,
      completedDate,
      ganttStart,
      ganttEnd,
      viewStartDate,
      viewDueDate,
      viewCompletedDate,
      viewGanttStart,
      viewGanttEnd,
      late: /^true|yes|是|1$/i.test(lateRaw),
      checklist: checklistMeta.summary,
      checklistItems: checklistMeta.items,
      checklistEntries: checklistMeta.entries,
      checklistDoneItems: checklistMeta.doneItems,
      checklistDoneCount: checklistMeta.doneCount,
      checklistTotalCount: checklistMeta.totalCount,
      labels,
      description: notes,
      bucketWasMapped: bucketResult.mapped,
      userWasMapped: userResult.mapped,
      raw: row
    };
  }).filter(task => task.task && !isLikelyId(task.task));
}

function resolveBucket(value, context) {
  const raw = cleanText(value);
  if (!raw) return { value: "", mapped: false };
  if (context.buckets.has(raw)) return { value: context.buckets.get(raw), mapped: true };
  return { value: readableGroup(raw) || "", mapped: false };
}

function resolveUsers(value, context) {
  const raw = cleanText(value);
  if (!raw) return { value: "", mapped: false };
  const parts = splitPeopleRaw(raw);
  let mapped = false;
  const names = parts.map(part => {
    const clean = cleanText(part);
    const lookupKey = clean.toLowerCase();
    if (context.users.has(clean)) { mapped = true; return context.users.get(clean); }
    if (context.users.has(lookupKey)) { mapped = true; return context.users.get(lookupKey); }
    return prettifyUser(clean);
  }).filter(Boolean).filter(name => !isLikelyId(name));
  return { value: distinctPreserveOrder(names).join("; ") || "", mapped };
}

function splitPeopleRaw(value) {
  const text = cleanText(value);
  if (!text) return [];
  if (/[;；]/.test(text)) return text.split(/[;；]/).map(cleanText).filter(Boolean);
  const uuidParts = text.split(/[,\n\t ]+/).map(cleanText).filter(Boolean);
  if (uuidParts.length > 1 && uuidParts.every(isLikelyId)) return uuidParts;
  return [text];
}

function parseChecklistMeta(row) {
  const doneRaw = cleanText(pick(row, ["Completed Checklist Items", "Completed checklist items", "已完成检查项"]));
  const itemsRaw = cleanText(pick(row, ["Checklist Items", "Checklist items", "Checklist", "检查清单项目", "检查清单"]));
  const itemEntries = parseChecklistEntries(itemsRaw);
  const doneEntries = parseChecklistEntries(doneRaw);
  const items = itemEntries.map(item => item.text);
  const inferredDoneItems = itemEntries.filter(item => item.done).map(item => item.text);
  const doneItems = doneEntries.length ? doneEntries.map(item => item.text) : inferredDoneItems;
  const doneProgress = parseChecklistProgress(doneRaw);
  const itemsProgress = parseChecklistProgress(itemsRaw);
  const doneCount = doneProgress?.done ?? parseChecklistCount(doneRaw) ?? doneItems.length;
  const totalCount = itemsProgress?.total ?? doneProgress?.total ?? parseChecklistCount(itemsRaw) ?? items.length;
  const summaryParts = [];

  if (totalCount || doneCount) {
    const safeTotal = Math.max(totalCount, doneCount);
    summaryParts.push(`${Math.min(doneCount, safeTotal)}/${safeTotal}`);
  }
  if (items.length) summaryParts.push(items.slice(0, 3).join(" · "));
  else if (!summaryParts.length && itemsRaw) summaryParts.push(itemsRaw);

  return {
    summary: summaryParts.join(" · "),
    items,
    entries: itemEntries,
    doneItems,
    doneCount,
    totalCount
  };
}

function parseChecklistProgress(value) {
  const text = cleanText(value);
  if (!text) return null;
  const ratio = text.match(/(\d+)\s*\/\s*(\d+)/);
  if (!ratio) return null;
  return { done: Number(ratio[1]), total: Number(ratio[2]) };
}

function parseChecklistItems(value) {
  return parseChecklistEntries(value).map(item => item.text);
}

function parseChecklistEntries(value) {
  const text = cleanText(value);
  if (!text) return [];
  const normalized = text
    .replace(/\s*[|｜；;]\s*/g, "\n")
    .replace(/\s*[•▪▫]\s*/g, "\n")
    .replace(/\s*[,，](?=\s*[A-Za-z\u4e00-\u9fff0-9\[☐☑✓✔✗✘❌❎❏■□▢▣])/g, "\n");
  const parts = normalized.split(/\n+/).map(cleanText).filter(Boolean);
  if (parts.length === 1 && /^\d+(?:\s*\/\s*\d+)?$/.test(parts[0])) return [];
  return distinctPreserveOrder(parts.map(parseChecklistEntry).filter(item => item.text), item => item.text.toLowerCase());
}

function parseChecklistEntry(value) {
  const raw = cleanText(value);
  if (!raw) return { text: "", done: false };
  const doneMarker = /^(?:\[x\]|\[X\]|☑|✅|✔|✓|已完成[:：]?|done[:：]?|complete[:：]?|completed[:：]?|true[:：]?|1[:：])\s*/i;
  const openMarker = /^(?:\[\s*\]|☐|⬜|□|未完成[:：]?|todo[:：]?|false[:：]?|0[:：])\s*/i;
  let done = false;
  let text = raw;

  if (doneMarker.test(text)) {
    done = true;
    text = text.replace(doneMarker, "");
  } else if (openMarker.test(text)) {
    text = text.replace(openMarker, "");
  }

  text = text
    .replace(/^[-*]\s*/, "")
    .replace(/^\d+[.)、]\s*/, "")
    .replace(/^第\s*\d+\s*項\s*/, "")
    .replace(/^checklist\s*item\s*[:：]?/i, "")
    .trim();

  return { text, done };
}

function parseChecklistCount(value) {
  const text = cleanText(value);
  if (!text) return null;
  if (/^\d+$/.test(text)) return Number(text);
  const ratio = text.match(/(\d+)\s*\/\s*(\d+)/);
  if (ratio) return Number(ratio[2]);
  return null;
}

function pick(row, names) {
  const keys = Object.keys(row || {});
  for (const name of names) {
    const key = keys.find(k => normalizeHeader(k) === normalizeHeader(name));
    if (key && row[key] !== undefined && row[key] !== null && String(row[key]).trim() !== "") return row[key];
  }
  return "";
}

function normalizeHeader(value) {
  return String(value || "").trim().toLowerCase().replace(/[\s_\-:：/\\()[\]（）]/g, "");
}

function cleanText(value) {
  if (value === undefined || value === null) return "";
  return String(value).replace(/\s+/g, " ").trim();
}

function readableProjectName(value, index) {
  const text = cleanText(value);
  if (!text || isLikelyId(text)) return `${t("unnamed")} ${index + 1}`;
  return text;
}

function readableGroup(value) {
  const text = cleanText(value);
  if (!text || isLikelyId(text)) return "";
  return text;
}

function prettifyUser(value) {
  let text = cleanText(value);
  if (!text) return "";
  const displayWithEmail = text.match(/^(.+?)\s*<[^>]+>$/);
  if (displayWithEmail) text = displayWithEmail[1];
  if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(text)) {
    text = text.split("@")[0].replace(/[._-]+/g, " ");
  }
  text = text.replace(/^i:0[#.|].*?\|/i, "").replace(/^.*?\\/, "").replace(/\s*\([^)]*@[^)]*\)\s*/g, "").replace(/\s+/g, " ").trim();
  if (isLikelyId(text)) return "";
  return titleCaseLoose(text);
}

function isLikelyId(value) {
  const text = cleanText(value);
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(text) ||
    /^[0-9a-f]{24,}$/i.test(text) ||
    /^[-_A-Za-z0-9]{18,}$/.test(text) && /\d/.test(text) && /[A-Za-z]/.test(text) && !/[\s,]/.test(text) ||
    /^\d{8,}$/.test(text);
}

function titleCaseLoose(value) {
  if (/[\u4e00-\u9fff]/.test(value)) return value;
  return value.replace(/\b[a-z]/g, ch => ch.toUpperCase());
}

function parseDate(value) {
  if (value === undefined || value === null || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return startOfDay(value);
  if (typeof value === "number") {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    return startOfDay(new Date(excelEpoch.getTime() + value * 86400000));
  }
  const text = String(value).trim();
  if (!text || ["na", "n/a", "none", "null"].includes(text.toLowerCase())) return null;
  const normalized = text.replace(/\./g, "/").replace(/年|月/g, "/").replace(/日/g, "").replace(/\s+\d{1,2}:\d{2}(:\d{2})?.*$/, "");
  const direct = new Date(normalized);
  if (!Number.isNaN(direct.getTime())) return startOfDay(direct);
  const parts = normalized.split(/[/-]/).map(x => Number(x));
  if (parts.length >= 3 && parts.every(Number.isFinite)) {
    const [a, b, c] = parts;
    if (a > 1900) return startOfDay(new Date(a, b - 1, c));
    if (c > 1900) return startOfDay(new Date(c, a - 1, b));
  }
  return null;
}

function inferStatus(progress, dueDate, completedDate) {
  const lower = String(progress || "").toLowerCase();
  const today = getProjectToday();
  const dueForStatus = toProjectYear(dueDate);
  if (completedDate || /complete|completed|已完成|完成/.test(lower)) return "Completed";
  if (dueForStatus && dueForStatus < today) return "Overdue";
  if (/in progress|进行中|进行/.test(lower)) return "In progress";
  if (/not started|未开始|未启动/.test(lower)) return "Not started";
  if (dueForStatus && dueForStatus <= addDays(today, 7)) return "Due soon";
  return progress || "Unknown";
}

function initializeFilters() {
  const tasks = state.tasks;
  fillFilterList("bucketFilter", unique(tasks.map(t => t.bucket)), "bucketCount");
  fillFilterList("assigneeFilter", unique(splitAssignees(tasks).map(t => t.assignedTo)), "assigneeCount");
  fillFilterList("statusFilter", unique(tasks.map(t => t.status)), "statusCount");

  const scheduledDates = tasks.flatMap(t => [t.viewGanttStart, t.viewGanttEnd, t.viewDueDate, t.viewCompletedDate]).filter(Boolean);
  if (scheduledDates.length) scheduledDates.push(getProjectToday());
  if (scheduledDates.length) {
    els.startDate.value = formatInputDate(new Date(Math.min(...scheduledDates)));
    els.endDate.value = formatInputDate(new Date(Math.max(...scheduledDates)));
  } else {
    els.startDate.value = "";
    els.endDate.value = "";
  }
}

function fillSelect(select, values, selectedValue = null) {
  select.innerHTML = "";
  values.forEach(value => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    if (value === selectedValue) option.selected = true;
    select.appendChild(option);
  });
}

function fillFilterList(id, values, countId) {
  const container = els[id];
  container.innerHTML = "";
  values.forEach(value => {
    const label = document.createElement("label");
    label.className = "option-pill";
    label.innerHTML = `<input type="checkbox" value="${escapeAttr(value)}" checked><span>${escapeHtml(value)}</span>`;
    container.appendChild(label);
  });
  refreshFilterCount(id, countId);
}

function selectedValues(container) {
  return Array.from(container.querySelectorAll("input[type='checkbox']:checked")).map(input => input.value);
}

function refreshFilterCount(id, countId) {
  const inputs = Array.from(els[id].querySelectorAll("input[type='checkbox']"));
  const checked = inputs.filter(input => input.checked).length;
  if (els[countId]) els[countId].textContent = checked === inputs.length ? String(inputs.length) : `${checked}/${inputs.length}`;
}

function resetFilters() {
  if (!state.tasks.length) return;
  initializeFilters();
  els.searchInput.value = "";
  applyFiltersAndRender();
}

function applyFiltersAndRender() {
  if (!state.tasks.length) { clearVisuals(); return; }

  const buckets = selectedValues(els.bucketFilter);
  const assignees = selectedValues(els.assigneeFilter);
  const statuses = selectedValues(els.statusFilter);
  const hasBucketOptions = hasFilterOptions(els.bucketFilter);
  const hasAssigneeOptions = hasFilterOptions(els.assigneeFilter);
  const hasStatusOptions = hasFilterOptions(els.statusFilter);
  refreshFilterCount("bucketFilter", "bucketCount");
  refreshFilterCount("assigneeFilter", "assigneeCount");
  refreshFilterCount("statusFilter", "statusCount");

  const start = parseDate(els.startDate.value);
  const end = parseDate(els.endDate.value);
  const query = els.searchInput.value.trim().toLowerCase();

  state.filtered = state.tasks.filter(task => {
    const taskAssignees = getTaskAssignees(task);
    const inAssignee = !hasAssigneeOptions || taskAssignees.some(a => assignees.includes(a)) || assignees.includes(task.assignedTo);
    const hasSchedule = task.viewGanttStart && task.viewGanttEnd;
    const inDate = !start || !end || !hasSchedule || (task.viewGanttStart <= end && task.viewGanttEnd >= start);
    const inSearch = !query || [
      task.task, task.bucket, task.assignedTo, task.status, task.progress, task.priority,
      task.checklist, task.labels, task.description, task.createdBy
    ].join(" ").toLowerCase().includes(query);
    return (!hasBucketOptions || buckets.includes(task.bucket)) && inAssignee &&
      (!hasStatusOptions || statuses.includes(task.status)) && inDate && inSearch;
  });

  renderAll();
}

function hasFilterOptions(container) {
  return container.querySelectorAll("input[type='checkbox']").length > 0;
}

function renderAll() {
  updateStats();
  renderGantt();
  renderAnalytics();
  renderPMO();
  renderTable();
}

function toggleFilterPanel() {
  state.filtersCollapsed = !state.filtersCollapsed;
  document.body.classList.toggle("filters-collapsed", state.filtersCollapsed);
  syncFilterPanelToggle();
  if (state.filtered.length) renderGantt({ viewport: captureGanttViewport() });
}

function syncFilterPanelToggle() {
  if (!els.filterCollapseBtn) return;
  const collapsed = state.filtersCollapsed;
  const label = collapsed ? t("filterExpand") : t("filterCollapse");
  els.filterCollapseBtn.textContent = label;
  els.filterCollapseBtn.setAttribute("aria-expanded", String(!collapsed));
  els.filterCollapseBtn.setAttribute("aria-label", label);
}

function updateStats() {
  const tasks = state.filtered;
  const total = tasks.length;
  const completed = tasks.filter(t => t.status === "Completed").length;
  const overdue = tasks.filter(t => t.status === "Overdue" || (t.viewDueDate && t.viewDueDate < getProjectToday() && t.status !== "Completed")).length;
  const noDue = tasks.filter(t => !t.viewDueDate).length;

  els.totalTasks.textContent = total.toLocaleString();
  els.completionRate.textContent = total ? `${Math.round(completed / total * 100)}%` : "0%";
  els.completedNote.textContent = t("completedNote").replace("{n}", completed);
  els.overdueTasks.textContent = overdue.toLocaleString();
  els.noDueTasks.textContent = noDue.toLocaleString();
}

function clearVisuals() {
  state.filtered = [];
  updateStats();
  if (els.ganttChart) els.ganttChart.innerHTML = `<div class="empty">${escapeHtml(t("emptyGantt"))}</div>`;
  const priorityChart = document.getElementById("priorityChart");
  const assigneeCompletionChart = document.getElementById("assigneeCompletionChart");
  const weeklyDueChart = document.getElementById("weeklyDueChart");
  if (priorityChart) priorityChart.innerHTML = `<div class="empty small">${escapeHtml(t("noData"))}</div>`;
  if (assigneeCompletionChart) assigneeCompletionChart.innerHTML = `<div class="empty small">${escapeHtml(t("noData"))}</div>`;
  if (weeklyDueChart) weeklyDueChart.innerHTML = `<div class="empty small">${escapeHtml(t("noData"))}</div>`;
  if (els.healthPanel) els.healthPanel.innerHTML = `<div class="empty small">${escapeHtml(t("noData"))}</div>`;
  if (els.workloadPanel) els.workloadPanel.innerHTML = `<div class="empty small">${escapeHtml(t("noData"))}</div>`;
  if (els.bucketPanel) els.bucketPanel.innerHTML = `<div class="empty small">${escapeHtml(t("noData"))}</div>`;
  if (els.riskPanel) els.riskPanel.innerHTML = `<div class="empty small">${escapeHtml(t("noRisk"))}</div>`;
  document.querySelector("#taskTable tbody").innerHTML = "";
  if (els.tableCount) els.tableCount.textContent = t("taskCount").replace("{n}", "0");
}

function renderGantt(options = {}) {
  const tasks = state.filtered.slice();
  const scheduled = tasks.filter(t => t.viewDueDate && t.viewGanttStart && t.viewGanttEnd);
  const el = els.ganttChart;
  const viewport = options.viewport || captureGanttViewport();

  if (!tasks.length) {
    el.innerHTML = `<div class="empty">${escapeHtml(t("emptyFilter"))}</div>`;
    return;
  }

  const bounds = getTimelineBounds(scheduled);
  const timelineStart = bounds.start;
  const timelineEnd = bounds.end;
  const totalDays = Math.max(1, dayDiff(timelineStart, addDays(timelineEnd, 1)));
  const pxPerDay = totalDays > 420 ? 4 : totalDays > 240 ? 5.5 : totalDays > 120 ? 7 : totalDays > 70 ? 9 : totalDays > 35 ? 12 : 16;
  const leftWidth = totalGanttLeftWidth();
  const availableTimeline = Math.max(720, (el.clientWidth || 1500) - leftWidth - 26);
  const timelineWidth = Math.max(availableTimeline, Math.round(totalDays * pxPerDay));
  // Grid lines: show daily lines if there's enough space, otherwise weekly
  const weekWidth = pxPerDay >= 8 ? pxPerDay : pxPerDay * 7;
  const today = getProjectToday();
  const todayLeft = dayDiff(timelineStart, today) * pxPerDay;
  const groupField = els.groupBy.value || "bucket";
  const subGroupField = els.subGroupBy?.value || "none";
  const renderTasks = expandTasksForGrouping(tasks, [groupField, subGroupField]).sort(compareForGantt);

  const dayTicks = buildDayTicks(timelineStart, timelineEnd, pxPerDay);
  const months = buildMonthBands(timelineStart, timelineEnd, pxPerDay);
  const monthHtml = months.map(month => `
    <div class="month-band" style="left:${month.left}px;width:${month.width}px">${escapeHtml(month.label)}</div>
  `).join("");
  const dayTicksHtml = dayTicks.map(tick => {
    const left = tick.left;
    const width = tick.width;
    return `<div class="day-cell ${tick.isToday ? "is-today" : ""}" style="left:${left}px;width:${width}px" data-date="${tick.fullDate}"><span>${tick.day}</span></div>`;
  }).join("");
  const todayHtml = today >= timelineStart && today <= timelineEnd
    ? `<div class="today-line project-style" style="left:${todayLeft}px"><span>${escapeHtml(t("todayLabel"))}</span></div>`
    : "";

  const groups = groupTasks(renderTasks, task => getGroupLabel(task, groupField));
  const plannedPercent = tasks.length ? Math.round(scheduled.length / tasks.length * 100) : 0;
  const summaryHtml = `
    <div class="project-summary-line">
      <strong>${PROJECT_YEAR}</strong> ${t("ganttBaseYear")} · ${tasks.length} ${t("ganttTasks")} · ${scheduled.length} ${t("ganttScheduled")} · ${plannedPercent}% ${t("ganttCoverage")} · ${Object.keys(groups).length} ${t("ganttSections")}
      <span>${t("ganttResizeHint")}</span>
    </div>
  `;

  const bodyTodayHtml = today >= timelineStart && today <= timelineEnd
    ? `<div class="project-today-body" style="left:calc(var(--left) + ${todayLeft}px)"></div>`
    : "";

  const groupHtml = Object.entries(groups).map(([groupName, groupRows], index) => {
    const groupScheduled = groupRows.filter(t => t.viewDueDate && t.viewGanttStart && t.viewGanttEnd);
    const groupRange = summarizeRange(groupScheduled);
    const statusCounts = countBy(groupRows, t => t.status);
    const groupBar = renderGroupRangeBar(groupScheduled, timelineStart, totalDays, pxPerDay);

    let innerHtml;
    if (subGroupField !== "none") {
      const subGroups = groupTasks(groupRows, task => getGroupLabel(task, subGroupField));
      innerHtml = Object.entries(subGroups).map(([subName, subRows]) => {
        const subRowsHtml = subRows.map((task, taskIndex) => renderGanttRow(task, timelineStart, totalDays, pxPerDay, timelineWidth, taskIndex)).join("");
        return `
          <details class="gantt-subgroup" open style="--timeline-width:${timelineWidth}px">
            <summary class="gantt-grid-row subgroup-row" style="--timeline-width:${timelineWidth}px">
              <div class="project-subgroup-left sticky-left">
                <span class="project-twisty small" aria-hidden="true"></span>
                <span class="project-subgroup-type">${escapeHtml(groupFieldCaption(subGroupField))}</span>
                <span class="subgroup-label">${escapeHtml(subName)}</span>
                <span class="subgroup-count">${subRows.length}</span>
              </div>
              <div class="project-subgroup-lane" style="width:${timelineWidth}px;--week-width:${weekWidth}px"></div>
            </summary>
            <div class="project-task-list">${subRowsHtml}</div>
          </details>
        `;
      }).join("");
    } else {
      innerHtml = `<div class="project-task-list">${groupRows.map((task, taskIndex) => renderGanttRow(task, timelineStart, totalDays, pxPerDay, timelineWidth, taskIndex)).join("")}</div>`;
    }

    return `
      <details class="gantt-group project-group ${index % 2 ? "alt" : ""}" open style="--timeline-width:${timelineWidth}px;--week-width:${weekWidth}px">
        <summary class="gantt-grid-row group-row" style="--timeline-width:${timelineWidth}px">
          <div class="project-group-left sticky-left">
            <div class="project-group-title-cell">
              <span class="project-twisty" aria-hidden="true"></span>
              <span class="project-section-label">SECTION</span>
              <span class="project-group-name">${escapeHtml(groupField === "none" ? t("allTasks") : groupName)}</span>
            </div>
            <div class="project-group-owner-cell">
              <span class="project-group-count">${groupRows.length}</span>
            </div>
            <div class="project-group-schedule-cell">
              <span class="project-group-range">${escapeHtml(groupRange)}</span>
            </div>
          </div>
          <div class="project-group-lane" style="width:${timelineWidth}px;--week-width:${weekWidth}px">
            ${groupBar}
            <div class="group-badges project-badges">${renderMiniBadges(statusCounts, groupRows)}</div>
          </div>
        </summary>
        ${innerHtml}
      </details>
    `;
  }).join("");

  el.innerHTML = `
    <div class="gantt-board project-gantt" data-timeline-width="${timelineWidth}" style="${ganttColumnStyle(timelineWidth)}">
      ${summaryHtml}
      <div class="gantt-scroll project-scroll">
        <div class="gantt-header project-header gantt-grid-row" style="--timeline-width:${timelineWidth}px">
          <div class="project-table-head sticky-left">
            <div class="head-cell">Task Name<span class="col-resizer" data-col="task" title="Resize"></span></div>
            <div class="head-cell">Owner<span class="col-resizer" data-col="owner" title="Resize"></span></div>
            <div class="head-cell">${state.lang === "zh" ? "起止日期" : "Schedule"}<span class="col-resizer" data-col="schedule" title="Resize"></span></div>
          </div>
          <div class="project-timeline-head" style="width:${timelineWidth}px;--week-width:${weekWidth}px" data-px-per-day="${pxPerDay}" data-timeline-start="${timelineStart.toISOString()}">
            <div class="month-layer">${monthHtml}</div>
            <div class="day-layer">${dayTicksHtml}</div>
            ${todayHtml}
            <div class="hover-date-indicator" style="display:none"></div>
          </div>
        </div>
        <div class="gantt-body project-body">${bodyTodayHtml}${groupHtml}</div>
      </div>
    </div>
  `;
  state.ganttTaskLookup = new Map(renderTasks.map(task => [task.renderKey, task]));
  bindGanttColumnResizers();
  bindGanttScrollControls();
  bindTaskTooltips();
  restoreGanttViewport(viewport, { timelineStart, timelineEnd, pxPerDay, timelineWidth, leftWidth });
}

function renderGroupRangeBar(rows, timelineStart, totalDays, pxPerDay) {
  if (!rows.length) return "";
  const minStart = new Date(Math.min(...rows.map(t => t.viewGanttStart)));
  const maxEnd = new Date(Math.max(...rows.map(t => t.viewGanttEnd)));
  const startOffset = clamp(dayDiff(timelineStart, minStart), 0, totalDays);
  const endOffset = clamp(dayDiff(timelineStart, addDays(maxEnd, 1)), 0, totalDays);
  const left = startOffset * pxPerDay;
  const width = Math.max(12, (endOffset - startOffset) * pxPerDay);
  return `<div class="project-summary-bar" style="left:${left}px;width:${width}px"></div>`;
}

function renderGanttRow(task, timelineStart, totalDays, pxPerDay, timelineWidth, taskIndex) {
  const hasBar = Boolean(task.viewDueDate && task.viewGanttStart && task.viewGanttEnd);
  const color = palette[task.priority] || palette.Normal;
  const dueClass = task.status === "Overdue" ? "danger" : isDueSoon(task) ? "warning" : "";
  const rowNumber = String(taskIndex + 1).padStart(2, "0");
  const assigneeText = formatTaskAssigneeDisplay(task);

  let laneContent = `<span class="project-unscheduled">No due date</span>`;
  if (hasBar) {
    const endForBar = task.viewGanttEnd <= task.viewGanttStart ? addDays(task.viewGanttStart, 1) : task.viewGanttEnd;
    const startOffset = clamp(dayDiff(timelineStart, task.viewGanttStart), 0, totalDays);
    const endOffset = clamp(dayDiff(timelineStart, addDays(endForBar, 1)), 0, totalDays);
    const left = startOffset * pxPerDay;
    const width = Math.max(12, (endOffset - startOffset) * pxPerDay);
    const duration = Math.max(1, dayDiff(task.viewGanttStart, endForBar) + 1);
    laneContent = `
      <div class="project-task-bar ${dueClass}" data-task-key="${escapeAttr(task.renderKey)}" data-bar-left="${left}" data-bar-width="${width}" style="left:${left}px;width:${width}px;background:${color}">
        <span>${duration}d</span>
      </div>
    `;
  }

  return `
    <div class="gantt-grid-row project-task-row ${hasBar ? "" : "no-bar"}" data-task-key="${escapeAttr(task.renderKey)}" style="--timeline-width:${timelineWidth}px">
      <div class="project-table-row sticky-left">
        <div class="project-task-name"><span class="row-index">${rowNumber}</span><span>${escapeHtml(task.task)}</span></div>
        <div class="project-owner">${escapeHtml(assigneeText)}</div>
        ${buildScheduleCellHtml(task)}
      </div>
      <div class="project-lane" style="width:${timelineWidth}px;--week-width:${pxPerDay * 7}px">${laneContent}</div>
    </div>
  `;
}

function buildMonthBands(start, end, pxPerDay) {
  const bands = [];
  const cursor = new Date(start.getFullYear(), start.getMonth(), 1);
  while (cursor <= end) {
    const monthStart = cursor < start ? start : new Date(cursor);
    const nextMonth = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
    const monthEnd = nextMonth > addDays(end, 1) ? addDays(end, 1) : nextMonth;
    const left = Math.max(0, dayDiff(start, monthStart) * pxPerDay);
    const width = Math.max(42, dayDiff(monthStart, monthEnd) * pxPerDay);
    bands.push({ left, width, label: formatMonth(cursor) });
    cursor.setMonth(cursor.getMonth() + 1);
  }
  return bands;
}

function getTimelineBounds(scheduled) {
  const startInput = parseDate(els.startDate.value);
  const endInput = parseDate(els.endDate.value);
  if (startInput && endInput) return { start: startInput, end: endInput };
  if (scheduled.length) {
    return {
      start: new Date(Math.min(...scheduled.map(t => t.viewGanttStart))),
      end: new Date(Math.max(...scheduled.map(t => t.viewGanttEnd)))
    };
  }
  const today = getProjectToday();
  return { start: startOfWeek(today), end: addDays(startOfWeek(today), 28) };
}

function priorityWeight(priority) {
  const p = (priority || "").toLowerCase();
  if (p === "urgent") return 0;
  if (p === "important") return 1;
  if (p === "medium") return 2;
  if (p === "normal") return 3;
  if (p === "low") return 4;
  return 3;
}

function compareForGantt(a, b) {
  const groupField = els.groupBy?.value || "bucket";
  return getGroupLabel(a, groupField).localeCompare(getGroupLabel(b, groupField), "zh-CN") ||
    priorityWeight(a.priority) - priorityWeight(b.priority) ||
    Number(!a.viewDueDate) - Number(!b.viewDueDate) ||
    (a.viewDueDate || new Date(8640000000000000)) - (b.viewDueDate || new Date(8640000000000000)) ||
    a.task.localeCompare(b.task, "zh-CN");
}

function groupTasks(items, getKey) {
  return items.reduce((acc, item) => {
    const key = getKey(item) || (state.lang === "zh" ? "未分類" : "Uncategorized");
    acc[key] ||= [];
    acc[key].push(item);
    return acc;
  }, {});
}

function getGroupLabel(task, field) {
  if (field === "none") return t("allTasks");
  if (field === "dueWeek") {
    if (!task.viewDueDate) return t("unscheduled");
    return formatWeekRangeLabel(task.viewDueDate);
  }
  const labelMap = { bucket: t("ungrouped"), assignedTo: t("unassigned"), status: t("unknown"), priority: "Normal" };
  const raw = cleanText(task.groupValues?.[field] ?? task[field]);
  if (!raw || isLikelyId(raw)) return labelMap[field] || t("ungrouped");
  return raw;
}

function groupFieldCaption(field) {
  if (field === "assignedTo") return state.lang === "zh" ? "負責人" : "Owner";
  if (field === "bucket") return state.lang === "zh" ? "分組" : "Group";
  if (field === "status") return state.lang === "zh" ? "狀態" : "Status";
  if (field === "priority") return state.lang === "zh" ? "優先級" : "Priority";
  if (field === "dueWeek") return state.lang === "zh" ? "截止週" : "Due Week";
  return state.lang === "zh" ? "分類" : "Group";
}

function summarizeRange(rows) {
  if (!rows.length) return t("unscheduled");
  const minStart = new Date(Math.min(...rows.map(t => t.viewGanttStart)));
  const maxEnd = new Date(Math.max(...rows.map(t => t.viewGanttEnd)));
  return `${formatDate(minStart)} → ${formatDate(maxEnd)}`;
}

function renderMiniBadges(statusCounts, rows) {
  const overdue = statusCounts.Overdue || 0;
  const done = statusCounts.Completed || 0;
  const noDue = rows.filter(t => !t.viewDueDate).length;
  return [
    overdue ? `<span class="mini-badge danger">${t("overdue")} ${overdue}</span>` : "",
    done ? `<span class="mini-badge ok">${state.lang === "zh" ? "完成" : "Done"} ${done}</span>` : "",
    noDue ? `<span class="mini-badge muted">${t("unscheduled")} ${noDue}</span>` : ""
  ].filter(Boolean).join("");
}

function setGanttOpen(open) {
  els.ganttChart.querySelectorAll("details").forEach(detail => { detail.open = open; });
}

function totalGanttLeftWidth() {
  const c = state.ganttCols;
  return c.task + c.owner + c.schedule;
}

function ganttColumnStyle(timelineWidth = 900) {
  const c = state.ganttCols;
  return [
    `--task-col:${c.task}px`, `--owner-col:${c.owner}px`,
    `--schedule-col:${c.schedule}px`,
    `--left:${totalGanttLeftWidth()}px`, `--timeline-width:${timelineWidth}px`,
    `--board-width:${totalGanttLeftWidth() + timelineWidth}px`
  ].join(";");
}

function applyGanttColumnVars() {
  const root = els.ganttChart?.querySelector(".project-gantt");
  if (!root) return;
  const c = state.ganttCols;
  root.style.setProperty("--task-col", `${c.task}px`);
  root.style.setProperty("--owner-col", `${c.owner}px`);
  root.style.setProperty("--schedule-col", `${c.schedule}px`);
  root.style.setProperty("--left", `${totalGanttLeftWidth()}px`);
}

function bindGanttColumnResizers() {
  els.ganttChart.querySelectorAll(".col-resizer").forEach(handle => {
    handle.addEventListener("pointerdown", event => {
      event.preventDefault();
      event.stopPropagation();
      const col = handle.dataset.col;
      const startX = event.clientX;
      const startWidth = state.ganttCols[col];
      const min = { task: 220, owner: 104, schedule: 136 }[col] || 72;
      const max = { task: 680, owner: 280, schedule: 240 }[col] || 260;
      document.body.classList.add("resizing-gantt-col");
      handle.setPointerCapture?.(event.pointerId);
      const onMove = moveEvent => {
        state.ganttCols[col] = Math.round(clamp(startWidth + moveEvent.clientX - startX, min, max));
        applyGanttColumnVars();
      };
      const onUp = () => {
        document.body.classList.remove("resizing-gantt-col");
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
        renderGantt({ viewport: captureGanttViewport({ keepTimelineOrigin: true }) });
      };
      window.addEventListener("pointermove", onMove);
      window.addEventListener("pointerup", onUp, { once: true });
    });
  });
}

function captureGanttViewport(options = {}) {
  const scroller = els.ganttChart?.querySelector(".project-scroll");
  if (!scroller) return null;
  return {
    scrollLeft: scroller.scrollLeft,
    scrollTop: scroller.scrollTop,
    leftWidth: totalGanttLeftWidth(),
    keepTimelineOrigin: Boolean(options.keepTimelineOrigin)
  };
}

function restoreGanttViewport(viewport, context) {
  const scroller = els.ganttChart?.querySelector(".project-scroll");
  if (!scroller) return;

  const { timelineStart, timelineEnd, pxPerDay, leftWidth } = context;
  const maxScrollLeft = Math.max(0, scroller.scrollWidth - scroller.clientWidth);
  const maxScrollTop = Math.max(0, scroller.scrollHeight - scroller.clientHeight);

  if (!viewport) {
    alignGanttToToday({ timelineStart, timelineEnd, pxPerDay, leftWidth, scroller });
    return;
  }

  const nextLeft = viewport.keepTimelineOrigin
    ? viewport.scrollLeft + (leftWidth - viewport.leftWidth)
    : viewport.scrollLeft;

  scroller.scrollLeft = clamp(nextLeft, 0, maxScrollLeft);
  scroller.scrollTop = clamp(viewport.scrollTop, 0, maxScrollTop);
}

function alignGanttToToday({ timelineStart, timelineEnd, pxPerDay, leftWidth, scroller }) {
  const today = getProjectToday();
  if (!(today >= timelineStart && today <= timelineEnd)) {
    scroller.scrollLeft = clamp(leftWidth, 0, Math.max(0, scroller.scrollWidth - scroller.clientWidth));
    return;
  }
  const todayLeft = dayDiff(timelineStart, today) * pxPerDay;
  const visibleTimelineWidth = Math.max(0, scroller.clientWidth - leftWidth);
  const targetLeft = leftWidth + todayLeft - Math.round(visibleTimelineWidth * 0.35);
  scroller.scrollLeft = clamp(targetLeft, 0, Math.max(0, scroller.scrollWidth - scroller.clientWidth));
}

function bindGanttScrollControls() {
  const scroller = els.ganttChart?.querySelector(".project-scroll");
  if (!scroller) return;
  scroller.addEventListener("wheel", event => {
    if (event.target.closest(".col-resizer")) return;
    const canScrollX = scroller.scrollWidth > scroller.clientWidth;
    const horizontalIntent = event.shiftKey || (Math.abs(event.deltaX) > Math.abs(event.deltaY) * 1.35 && Math.abs(event.deltaX) > 2);
    if (horizontalIntent && canScrollX) {
      scroller.scrollLeft += event.deltaX || event.deltaY;
      event.preventDefault();
    }
  }, { passive: false });

  const timelineHead = els.ganttChart?.querySelector(".project-timeline-head");
  if (!timelineHead) return;

  const pxPerDay = parseFloat(timelineHead.dataset.pxPerDay) || 10;
  const timelineStartStr = timelineHead.dataset.timelineStart;
  const timelineStartDate = timelineStartStr ? new Date(timelineStartStr) : null;
  if (!timelineStartDate) return;

  const indicator = timelineHead.querySelector(".hover-date-indicator");
  const body = els.ganttChart.querySelector(".project-body");
  const timelineWidth = parseFloat(els.ganttChart.querySelector(".project-gantt")?.dataset.timelineWidth || "0") || 0;

  // Column highlight that spans the full body
  let colHighlight = els.ganttChart.querySelector(".gantt-col-highlight");
  if (!colHighlight) {
    colHighlight = document.createElement("div");
    colHighlight.className = "gantt-col-highlight";
    if (body) body.appendChild(colHighlight);
  }

  const leftWidth = totalGanttLeftWidth();

  const hideHover = () => {
    if (indicator) indicator.style.display = "none";
    if (colHighlight) colHighlight.style.display = "none";
  };

  const handleMove = (e) => {
    const scrollerRect = scroller.getBoundingClientRect();
    const xInScroller = e.clientX - scrollerRect.left + scroller.scrollLeft;
    const xInTimeline = xInScroller - leftWidth;
    if (xInTimeline < 0) {
      hideHover();
      return;
    }
    const dayOffset = Math.floor(xInTimeline / pxPerDay);
    const maxDayOffset = Math.max(0, Math.floor(Math.max(timelineWidth - 1, 0) / pxPerDay));
    const safeDayOffset = clamp(dayOffset, 0, maxDayOffset);
    const hoverCenter = clamp((safeDayOffset * pxPerDay) + (pxPerDay / 2), 40, Math.max(40, timelineWidth - 40));
    const hoverDate = addDays(timelineStartDate, dayOffset);
    const dayLeft = safeDayOffset * pxPerDay;

    if (indicator) {
      indicator.style.display = "flex";
      indicator.style.left = `${hoverCenter}px`;
      indicator.textContent = formatDate(hoverDate);
    }
    if (colHighlight) {
      colHighlight.style.display = "block";
      colHighlight.style.left = `${leftWidth + dayLeft}px`;
      colHighlight.style.width = `${pxPerDay}px`;
    }
  };

  const hoverSurfaces = [timelineHead, body].filter(Boolean);
  hoverSurfaces.forEach(surface => {
    surface.addEventListener("mousemove", handleMove);
    surface.addEventListener("mouseleave", hideHover);
  });

  scroller.addEventListener("scroll", hideHover, { passive: true });
  scroller.addEventListener("mouseleave", hideHover);
}

function toggleFocusMode() {
  const enabled = document.body.classList.toggle("focus-gantt");
  els.focusModeBtn.textContent = enabled ? t("focusModeExit") : t("focusMode");
  if (enabled) {
    document.getElementById("tab-gantt")?.scrollIntoView({ block: "start", behavior: "smooth" });
  }
}

function toggleGanttDensity() {
  const enabled = document.body.classList.toggle("compact-gantt");
  els.densityBtn.textContent = enabled ? t("densityBtnAlt") : t("densityBtn");
}

function renderAnalytics() {
  renderPriorityChart();
  renderAssigneeCompletionChart();
  renderWeeklyDueChart();
}

function renderPriorityChart() {
  const el = document.getElementById("priorityChart");
  if (!el) return;
  const tasks = state.filtered;
  if (!tasks.length) { el.innerHTML = `<div class="empty small">${escapeHtml(t("noData"))}</div>`; return; }
  const priorities = ["Urgent", "Important", "Medium", "Normal", "Low"];
  const counts = {};
  tasks.forEach(t => { const p = t.priority || "Normal"; counts[p] = (counts[p] || 0) + 1; });
  const max = Math.max(...Object.values(counts), 1);
  const colors = { Urgent: "#D92D20", Important: "#F79009", Medium: "#0054A6", Normal: "#0054A6", Low: "#8EA4BA" };
  el.innerHTML = priorities.filter(p => counts[p]).map(p => `
    <div class="h-bar-row">
      <span class="h-bar-label">${p}</span>
      <div class="h-bar-track"><div class="h-bar-fill" style="width:${Math.round((counts[p]||0)/max*100)}%;background:${colors[p]||'#0054A6'}"></div></div>
      <span class="h-bar-value">${counts[p]||0}</span>
    </div>
  `).join("");
}

function renderAssigneeCompletionChart() {
  const el = document.getElementById("assigneeCompletionChart");
  if (!el) return;
  const expanded = splitAssignees(state.filtered);
  if (!expanded.length) { el.innerHTML = `<div class="empty small">${escapeHtml(t("noData"))}</div>`; return; }
  const byPerson = {};
  expanded.forEach(t => {
    if (!byPerson[t.assignedTo]) byPerson[t.assignedTo] = { total: 0, done: 0 };
    byPerson[t.assignedTo].total++;
    if (t.status === "Completed") byPerson[t.assignedTo].done++;
  });
  const sorted = Object.entries(byPerson)
    .map(([name, d]) => ({ name, ...d, rate: d.total ? Math.round(d.done/d.total*100) : 0 }))
    .sort((a,b) => b.total - a.total).slice(0, 10);
  el.innerHTML = sorted.map(row => `
    <div class="h-bar-row">
      <span class="h-bar-label">${escapeHtml(row.name)}</span>
      <div class="h-bar-track"><div class="h-bar-fill completion" style="width:${row.rate}%"></div></div>
      <span class="h-bar-value">${row.rate}% (${row.done}/${row.total})</span>
    </div>
  `).join("");
}

function renderWeeklyDueChart() {
  const el = document.getElementById("weeklyDueChart");
  if (!el) return;
  const tasks = state.filtered.filter(t => t.viewDueDate);
  if (!tasks.length) { el.innerHTML = `<div class="empty small">${escapeHtml(t("noData"))}</div>`; return; }
  const weeks = {};
  tasks.forEach(t => { const w = startOfWeek(t.viewDueDate); const key = formatInputDate(w); weeks[key] = (weeks[key] || 0) + 1; });
  const sorted = Object.entries(weeks).sort((a,b) => a[0].localeCompare(b[0]));
  const max = Math.max(...sorted.map(s => s[1]), 1);
  el.innerHTML = `
    <div class="weekly-chart">
      ${sorted.map(([week, count]) => `
        <div class="weekly-col">
          <div class="weekly-bar-wrap"><div class="weekly-bar" style="height:${Math.round(count/max*100)}%"></div></div>
          <span class="weekly-label">${escapeHtml(formatWeekRangeLabel(new Date(week), { compact: true }))}</span>
          <span class="weekly-count">${count}</span>
        </div>
      `).join("")}
    </div>
  `;
}

function renderPMO() {
  renderHealthPanel();
  renderWorkloadPanel();
  renderBucketPanel();
  renderRiskPanel();
}

function renderHealthPanel() {
  const tasks = state.filtered;
  if (!tasks.length) { els.healthPanel.innerHTML = `<div class="empty small">${escapeHtml(t("noData"))}</div>`; return; }
  const total = tasks.length;
  const completed = tasks.filter(t => t.status === "Completed").length;
  const overdue = tasks.filter(t => t.status === "Overdue").length;
  const dueSoon = tasks.filter(isDueSoon).length;
  const noDue = tasks.filter(t => !t.viewDueDate).length;
  const active = total - completed;
  const items = [
    { label: t("healthDone"), value: completed, color: "ok" },
    { label: t("healthActive"), value: active, color: "blue" },
    { label: t("health7d"), value: dueSoon, color: "warning" },
    { label: t("healthOverdue"), value: overdue, color: "danger" },
    { label: t("healthNoDue"), value: noDue, color: "muted" }
  ];
  els.healthPanel.innerHTML = `
    <div class="health-score">
      <div class="score-ring" style="--score:${total ? Math.round(completed / total * 100) : 0}">
        <strong>${total ? Math.round(completed / total * 100) : 0}%</strong>
        <span>${t("completionLabel")}</span>
      </div>
      <div class="score-copy">
        <h4>${state.context.planName ? escapeHtml(state.context.planName) : "Planner Portfolio"}</h4>
        <p>${t("healthCopy").replace("{total}", total).replace("{active}", active)}</p>
      </div>
    </div>
    <div class="health-metrics">
      ${items.map(item => `<div class="metric ${item.color}"><span>${item.label}</span><strong>${item.value}</strong></div>`).join("")}
    </div>
  `;
}

function renderWorkloadPanel() {
  const expanded = splitAssignees(state.filtered);
  const counts = topEntries(countBy(expanded, t => t.assignedTo), 12);
  if (!counts.length) { els.workloadPanel.innerHTML = `<div class="empty small">${escapeHtml(t("noData"))}</div>`; return; }
  const max = Math.max(...counts.map(r => r.value), 1);
  els.workloadPanel.innerHTML = counts.map(row => {
    const personRows = expanded.filter(t => t.assignedTo === row.key);
    const risk = personRows.filter(t => t.status === "Overdue" || isDueSoon(t)).length;
    return `
      <div class="rank-row">
        <div class="rank-label"><strong>${escapeHtml(row.key)}</strong><span>${row.value} ${t("riskItems")} ${risk}</span></div>
        <div class="rank-bar"><i style="width:${Math.round(row.value / max * 100)}%"></i></div>
      </div>
    `;
  }).join("");
}

function renderBucketPanel() {
  const buckets = Object.entries(groupTasks(state.filtered, t => t.bucket || t("ungrouped")))
    .map(([bucket, rows]) => ({ bucket, rows, total: rows.length, done: rows.filter(t => t.status === "Completed").length }))
    .sort((a, b) => b.total - a.total || a.bucket.localeCompare(b.bucket, "zh-CN"));
  if (!buckets.length) { els.bucketPanel.innerHTML = `<div class="empty small">${escapeHtml(t("noData"))}</div>`; return; }
  els.bucketPanel.innerHTML = buckets.slice(0, 12).map(row => {
    const percent = row.total ? Math.round(row.done / row.total * 100) : 0;
    return `
      <div class="rank-row">
        <div class="rank-label"><strong>${escapeHtml(row.bucket)}</strong><span>${row.done}/${row.total} ${state.lang === "zh" ? "完成" : "done"}</span></div>
        <div class="rank-bar"><i style="width:${percent}%"></i></div>
      </div>
    `;
  }).join("");
}

function renderRiskPanel() {
  const risks = state.filtered
    .filter(t => t.status === "Overdue" || isDueSoon(t) || !t.viewDueDate)
    .sort((a, b) => riskWeight(b) - riskWeight(a) || (a.viewDueDate || new Date(8640000000000000)) - (b.viewDueDate || new Date(8640000000000000)))
    .slice(0, 18);
  if (!risks.length) { els.riskPanel.innerHTML = `<div class="empty small">${escapeHtml(t("noRisk"))}</div>`; return; }
  els.riskPanel.innerHTML = risks.map(task => {
    const type = task.status === "Overdue" ? t("overdue") : !task.viewDueDate ? t("noDueDate") : t("dueSoon");
    const cls = task.status === "Overdue" ? "danger" : !task.viewDueDate ? "muted" : "warning";
    return `
      <article class="risk-item">
        <span class="risk-type ${cls}">${type}</span>
        <div>
          <strong>${escapeHtml(task.task)}</strong>
          <p>${escapeHtml(task.bucket)} · ${escapeHtml(task.assignedTo)} · ${task.viewDueDate ? formatDate(task.viewDueDate) : t("unscheduled")}</p>
        </div>
      </article>
    `;
  }).join("");
}

function renderTable() {
  const tbody = document.querySelector("#taskTable tbody");
  tbody.innerHTML = "";
  if (els.tableCount) els.tableCount.textContent = t("taskCount").replace("{n}", state.filtered.length);
  state.filtered.forEach(task => {
    const tr = document.createElement("tr");
    const cells = [
      task.task, task.bucket, task.assignedTo,
      statusLabel(task.status), task.priority,
      formatDate(task.viewStartDate), formatDate(task.viewDueDate),
      formatDate(task.viewCompletedDate),
      [task.labels, task.checklist].filter(Boolean).join(" · ")
    ];
    cells.forEach((value, idx) => {
      const td = document.createElement("td");
      if (idx === 3) {
        td.innerHTML = `<span class="status-chip ${statusClass(task.status)}">${escapeHtml(value || "")}</span>`;
      } else if (idx === 6 && !value) {
        td.innerHTML = `<span class="due-chip muted">${t("noDueDate")}</span>`;
      } else {
        td.textContent = value || "";
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}

function exportCsv() {
  if (!state.filtered.length) return;
  const rows = state.filtered.map(t => ({
    TaskName: t.task, Bucket: t.bucket, AssignedTo: t.assignedTo,
    Status: t.status, PlannerStatus: t.progress, Priority: t.priority,
    StartDate: formatDate(t.viewStartDate), DueDate: formatDate(t.viewDueDate),
    CompletedDate: formatDate(t.viewCompletedDate), CreatedBy: t.createdBy,
    Labels: t.labels, Checklist: t.checklist, Notes: t.description
  }));
  const headers = Object.keys(rows[0]);
  const csv = [headers.join(","), ...rows.map(row => headers.map(h => csvEscape(row[h])).join(","))].join("\n");
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `planner_pmo_${formatInputDate(new Date())}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

function csvEscape(value) {
  const text = value === null || value === undefined ? "" : String(value);
  return /[",\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function splitAssignees(tasks) {
  return tasks.flatMap(task => {
    const names = getTaskAssignees(task);
    return (names.length ? names : [t("unassigned")]).map(name => ({ ...task, assignedTo: name }));
  });
}

function splitPersonNames(text) {
  return cleanText(text).split(/\s*[;；]\s*/).map(cleanText).filter(Boolean);
}

function getTaskAssignees(task) {
  const names = Array.isArray(task?.assignees) ? task.assignees : splitPersonNames(task?.assignedTo || "");
  return names.length ? distinctPreserveOrder(names) : [];
}

function expandTasksForGrouping(tasks, fields) {
  const needsAssigneeExpansion = fields.includes("assignedTo");
  return tasks.flatMap(task => {
    if (!needsAssigneeExpansion) return [{ ...task, renderKey: `${task.rowIndex}:base` }];
    const assignees = getTaskAssignees(task);
    const values = assignees.length ? assignees : [t("unassigned")];
    return values.map((assignee, index) => ({
      ...task,
      groupValues: { ...(task.groupValues || {}), assignedTo: assignee },
      renderKey: `${task.rowIndex}:assignedTo:${assignee}:${index}`
    }));
  });
}

function formatTaskAssigneeDisplay(task) {
  const assignees = getTaskAssignees(task);
  if (!assignees.length) return t("unassigned");
  const activeAssignee = cleanText(task.groupValues?.assignedTo);
  if (!activeAssignee || assignees.length === 1) return assignees.join("; ");
  const others = assignees.filter(name => name !== activeAssignee);
  return others.length ? `${activeAssignee} +${others.length}` : activeAssignee;
}

function bindTaskTooltips() {
  const rows = els.ganttChart?.querySelectorAll(".project-task-row[data-task-key]");
  if (!rows?.length) return;
  const tooltip = ensureTaskTooltip();
  const hoverDelay = 600;
  let pendingPoint = null;

  const clearPendingShow = () => {
    if (tooltip._showTimer) {
      window.clearTimeout(tooltip._showTimer);
      tooltip._showTimer = null;
    }
  };

  const show = (row, point) => {
    const task = state.ganttTaskLookup.get(row.dataset.taskKey);
    if (!task) return;
    tooltip.innerHTML = buildTaskTooltipHtml(task);
    tooltip.hidden = false;
    tooltip.classList.add("visible");
    positionTaskTooltip(point, tooltip);
  };

  const move = (event, row) => {
    pendingPoint = { clientX: event.clientX, clientY: event.clientY };
    if (tooltip.hidden) return;
    if (tooltip.dataset.taskKey !== row.dataset.taskKey) return;
    positionTaskTooltip(pendingPoint, tooltip);
  };

  const hide = () => {
    clearPendingShow();
    pendingPoint = null;
    tooltip.dataset.taskKey = "";
    tooltip.classList.remove("visible");
    tooltip.hidden = true;
  };

  const scheduleShow = (event, row) => {
    clearPendingShow();
    pendingPoint = { clientX: event.clientX, clientY: event.clientY };
    tooltip._showTimer = window.setTimeout(() => {
      tooltip._showTimer = null;
      tooltip.dataset.taskKey = row.dataset.taskKey;
      show(row, pendingPoint || { clientX: event.clientX, clientY: event.clientY });
    }, hoverDelay);
  };

  rows.forEach(row => {
    row.addEventListener("mouseenter", event => scheduleShow(event, row));
    row.addEventListener("mousemove", event => move(event, row));
    row.addEventListener("mouseleave", hide);
  });

  els.ganttChart.querySelector(".project-scroll")?.addEventListener("scroll", hide, { passive: true });
}

function ensureTaskTooltip() {
  let tooltip = document.getElementById("ganttTaskTooltip");
  if (!tooltip) {
    tooltip = document.createElement("div");
    tooltip.id = "ganttTaskTooltip";
    tooltip.className = "gantt-task-tooltip";
    tooltip.hidden = true;
    document.body.appendChild(tooltip);
  }
  return tooltip;
}

function positionTaskTooltip(event, tooltip) {
  const offsetX = 16;
  const offsetY = 16;
  const width = tooltip.offsetWidth || 360;
  const height = tooltip.offsetHeight || 180;
  const maxLeft = Math.max(12, window.innerWidth - width - 12);
  const left = clamp(event.clientX + offsetX, 12, maxLeft);
  const preferredTop = event.clientY - height - offsetY;
  const fallbackTop = event.clientY + offsetY;
  const top = preferredTop >= 12
    ? preferredTop
    : clamp(fallbackTop, 12, Math.max(12, window.innerHeight - height - 12));
  tooltip.style.left = `${left}px`;
  tooltip.style.top = `${top}px`;
}

function buildTaskTooltipHtml(task) {
  const assignees = getTaskAssignees(task);
  const startText = formatDate(task.viewStartDate) || (state.lang === "zh" ? "未設置" : "N/A");
  const dueText = formatDate(task.viewDueDate) || (state.lang === "zh" ? "未設置" : "N/A");
  const duration = task.viewGanttStart && task.viewGanttEnd
    ? Math.max(1, dayDiff(task.viewGanttStart, task.viewGanttEnd <= task.viewGanttStart ? addDays(task.viewGanttStart, 1) : task.viewGanttEnd) + 1)
    : null;
  const checklistProgress = task.checklistTotalCount
    ? `${Math.min(task.checklistDoneCount, task.checklistTotalCount)}/${task.checklistTotalCount}`
    : task.checklist;
  const checklistItems = buildChecklistDisplayItems(task);
  const assigneeDetail = assignees.join("; ") || t("unassigned");

  return `
    <div class="gantt-task-tooltip-card">
      <div class="gantt-task-tooltip-head">
        <strong>${escapeHtml(task.task)}</strong>
        <span class="status-chip ${statusClass(task.status)}">${escapeHtml(statusLabel(task.status))}</span>
      </div>
      <div class="gantt-task-tooltip-meta">
        <span>${escapeHtml(task.bucket)}</span>
        <span>${escapeHtml(task.priority || "Normal")}</span>
        ${duration ? `<span>${escapeHtml(String(duration))}${state.lang === "zh" ? " 天" : "d"}</span>` : ""}
      </div>
      <div class="gantt-task-tooltip-grid">
        <div class="gantt-task-tooltip-field">
          <label>${state.lang === "zh" ? "負責人" : "Owner"}</label>
          <p>${escapeHtml(assigneeDetail)}</p>
        </div>
        <div class="gantt-task-tooltip-field">
          <label>${state.lang === "zh" ? "建立者" : "Created by"}</label>
          <p>${escapeHtml(task.createdBy || "-")}</p>
        </div>
        <div class="gantt-task-tooltip-field gantt-task-tooltip-date-field">
          <label>${state.lang === "zh" ? "開始" : "Start"}</label>
          <p>${escapeHtml(startText)}</p>
        </div>
        <div class="gantt-task-tooltip-field gantt-task-tooltip-date-field">
          <label>${state.lang === "zh" ? "截止" : "Due"}</label>
          <p>${escapeHtml(dueText)}</p>
        </div>
      </div>
      ${task.labels ? `<div class="gantt-task-tooltip-block"><label>${state.lang === "zh" ? "標籤" : "Labels"}</label><p>${escapeHtml(task.labels)}</p></div>` : ""}
      ${task.description ? `<div class="gantt-task-tooltip-block"><label>${state.lang === "zh" ? "備註" : "Notes"}</label><p>${escapeHtml(task.description)}</p></div>` : ""}
      ${checklistProgress || checklistItems.length ? `
        <div class="gantt-task-tooltip-block">
          <label>${state.lang === "zh" ? "Checklist / 分任務" : "Checklist / Subtasks"}${checklistProgress ? `<span>${escapeHtml(checklistProgress)}</span>` : ""}</label>
          ${checklistItems.length ? `<ul>${checklistItems.slice(0, 8).map(item => `<li class="gantt-checklist-item ${item.done ? "done" : ""}"><span>${escapeHtml(item.text)}</span></li>`).join("")}</ul>` : `<p>${escapeHtml(task.checklist || (state.lang === "zh" ? "有清單資料" : "Checklist available"))}</p>`}
        </div>` : ""}
    </div>
  `;
}

function buildScheduleCellHtml(task) {
  const startText = formatDate(task.viewStartDate) || "—";
  const dueText = formatDate(task.viewDueDate) || (state.lang === "zh" ? "未設置" : "No date");
  return `
    <div class="project-schedule-cell ${task.viewDueDate ? "" : "is-open"}">
      <div class="project-schedule-item">
        <b>${state.lang === "zh" ? "開始：" : "Start:"}</b>
        <strong>${escapeHtml(startText)}</strong>
      </div>
      <div class="project-schedule-item ${task.viewDueDate ? "" : "is-open"}">
        <b>${state.lang === "zh" ? "截止：" : "Due:"}</b>
        <strong class="${task.viewDueDate ? "" : "muted-cell"}">${escapeHtml(dueText)}</strong>
      </div>
    </div>
  `;
}

function buildChecklistDisplayItems(task) {
  const entries = Array.isArray(task.checklistEntries) ? task.checklistEntries : [];
  const items = entries.length ? entries.map(item => item.text) : (Array.isArray(task.checklistItems) ? task.checklistItems : []);
  if (!items.length) return [];
  const doneItems = Array.isArray(task.checklistDoneItems) ? task.checklistDoneItems : [];
  const doneSet = new Set(doneItems.map(normalizeChecklistItemKey));
  const markAllDone = !doneSet.size && task.checklistTotalCount && task.checklistDoneCount >= task.checklistTotalCount;
  return items.map((item, index) => ({
    text: item,
    done: markAllDone || Boolean(entries[index]?.done) || doneSet.has(normalizeChecklistItemKey(item))
  }));
}

function normalizeChecklistItemKey(value) {
  return cleanText(value)
    .toLowerCase()
    .replace(/^(?:\[x\]|\[\s*\]|☑|✅|✔|✓|☐|⬜|□)\s*/i, "")
    .replace(/^[-*]\s*/, "")
    .replace(/^\d+[.)、]\s*/, "")
    .replace(/^第\s*\d+\s*項\s*/, "")
    .replace(/\s+/g, " ")
    .trim();
}

function countBy(items, getKey) {
  return items.reduce((acc, item) => { const key = getKey(item) || "Unknown"; acc[key] = (acc[key] || 0) + 1; return acc; }, {});
}

function topEntries(counts, limit) {
  return Object.entries(counts).map(([key, value]) => ({ key, value })).sort((a, b) => b.value - a.value || a.key.localeCompare(b.key, "zh-CN")).slice(0, limit);
}

function distinctPreserveOrder(values, keySelector = value => String(value)) {
  const seen = new Set();
  const results = [];
  values.filter(Boolean).forEach(value => {
    const key = keySelector(value);
    if (seen.has(key)) return;
    seen.add(key);
    results.push(value);
  });
  return results;
}

function unique(values) {
  return Array.from(new Set(values.filter(Boolean))).sort((a, b) => String(a).localeCompare(String(b), "zh-CN"));
}

function isDueSoon(task) {
  if (!task.viewDueDate || task.status === "Completed" || task.status === "Overdue") return false;
  const today = getProjectToday();
  return task.viewDueDate >= today && task.viewDueDate <= addDays(today, 7);
}

function riskWeight(task) {
  if (task.status === "Overdue") return 3;
  if (isDueSoon(task)) return 2;
  if (!task.viewDueDate) return 1;
  return 0;
}

function statusLabel(status) {
  const map = {
    Completed: t("statusCompleted"), Overdue: t("statusOverdue"),
    "Due soon": t("statusDueSoon"), "In progress": t("statusInProgress"),
    "Not started": t("statusNotStarted"), Unknown: t("statusUnknown")
  };
  return map[status] || status;
}

function statusClass(status) {
  return { Completed: "ok", Overdue: "danger", "Due soon": "warning", "In progress": "blue", "Not started": "muted", Unknown: "muted" }[status] || "muted";
}

/* ─── TIMELINE FIX: smarter tick labels that don't overlap ─── */
/* Build individual day cells for the timeline - like Google-style day numbers */
function buildDayTicks(start, end, pxPerDay) {
  const ticks = [];
  const totalDays = dayDiff(start, addDays(end, 1));
  const today = getProjectToday();

  // Decide which days to show labels for based on zoom
  // At very compressed zoom, skip days to avoid crowding
  let stepDays = 1;
  if (pxPerDay < 4) stepDays = 7;
  else if (pxPerDay < 6) stepDays = 4;
  else if (pxPerDay < 9) stepDays = 3;
  else if (pxPerDay < 14) stepDays = 2;

  const cursor = startOfDay(new Date(start));
  let dayIndex = 0;
  while (cursor <= end) {
    const showLabel = dayIndex % stepDays === 0;
    const left = dayIndex * pxPerDay;
    const width = pxPerDay * stepDays;
    const isToday = cursor.getTime() === today.getTime();
    if (showLabel) {
      ticks.push({
        left,
        width: Math.min(width, (totalDays - dayIndex) * pxPerDay),
        day: cursor.getDate(),
        fullDate: formatInputDate(cursor),
        isToday
      });
    }
    cursor.setDate(cursor.getDate() + stepDays);
    dayIndex += stepDays;
  }
  return ticks;
}

function startOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function toProjectYear(date) {
  if (!date) return null;
  const d = startOfDay(new Date(date));
  d.setFullYear(PROJECT_YEAR);
  return d;
}

function getProjectToday() {
  const now = startOfDay(new Date());
  now.setFullYear(PROJECT_YEAR);
  return now;
}

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function dayDiff(start, end) { return Math.ceil((end - start) / 86400000); }

function startOfWeek(date) {
  const d = startOfDay(date);
  const day = d.getDay() || 7;
  d.setDate(d.getDate() - day + 1);
  return d;
}

function formatWeekRangeLabel(date, options = {}) {
  const weekStart = startOfWeek(date);
  const weekEnd = addDays(weekStart, 6);
  if (options.compact) {
    return `${formatMonthDay(weekStart)}-${formatMonthDay(weekEnd)}`;
  }
  return `${formatInputDate(weekStart)} ~ ${formatInputDate(weekEnd)}`;
}

function formatMonthDay(date) {
  const d = new Date(date);
  return `${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")}`;
}

function formatDate(date) { return date ? formatInputDate(date) : ""; }

function formatInputDate(date) {
  if (!date) return "";
  const d = new Date(date);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function formatMonth(date) {
  const d = new Date(date);
  if (state.lang === "en") {
    const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    return monthNames[d.getMonth()];
  }
  const monthNames = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"];
  return monthNames[d.getMonth()];
}

function escapeHtml(value) {
  return String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function escapeAttr(value) { return escapeHtml(value).replace(/`/g, "&#96;"); }

function clamp(value, min, max) { return Math.min(max, Math.max(min, value)); }

function setNotice(message, type = "") {
  els.notice.textContent = message;
  els.notice.className = `notice ${type}`.trim();
}

let ganttResizeTimer = null;
window.addEventListener("resize", () => {
  if (!state.filtered.length || !els.ganttChart?.querySelector(".project-gantt")) return;
  clearTimeout(ganttResizeTimer);
  ganttResizeTimer = setTimeout(renderGantt, 120);
});
