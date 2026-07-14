"use strict";

const state = {
  bootstrap: null,
  sheetName: "",
  columns: [],
  rows: [],
  selected: new Set(),
  columnView: "student",
  search: "",
  searchScope: "current",
  searchResults: [],
  searchRequest: 0,
  searchTimer: null,
  quizNumber: "1",
  dirty: false,
  announcementDirty: false,
  busy: false,
};

const elements = {};

function byId(id) {
  return document.getElementById(id);
}

function rowKey(row) {
  return row.student_id ? `student:${row.student_id}` : `new:${row.client_id}`;
}

function fingerprint(row) {
  const values = row.values || {};
  return [
    String(values.uid || "").trim(),
    String(values["First Name"] || "").trim(),
    String(values["Last Name"] || "").trim(),
  ].join("|").toLowerCase();
}

function displayValue(value) {
  if (value === null || value === undefined) return "";
  if (value === true) return "TRUE";
  if (value === false) return "FALSE";
  return String(value);
}

async function api(path, options = {}) {
  const response = await fetch(path, {
    headers: { "Content-Type": "application/json", ...(options.headers || {}) },
    ...options,
  });
  let payload;
  try {
    payload = await response.json();
  } catch {
    payload = { ok: false, error: `The server returned HTTP ${response.status}.` };
  }
  if (!response.ok || payload.ok === false) {
    const error = new Error(payload.error || payload.output || "The action could not be completed.");
    error.payload = payload;
    throw error;
  }
  return payload;
}

function setSaveState(label, kind = "neutral") {
  elements.saveState.textContent = label;
  elements.saveState.className = `status-pill is-${kind}`;
}

function markDirty(kind = "sheet") {
  if (kind === "announcement") state.announcementDirty = true;
  else state.dirty = true;
  setSaveState("Unsaved changes", "dirty");
}

function setBusy(busy, title = "Working", detail = "Please keep this window open.") {
  state.busy = busy;
  elements.busyOverlay.hidden = !busy;
  elements.busyTitle.textContent = title;
  elements.busyDetail.textContent = detail;
  document.querySelectorAll("button, input, select, textarea").forEach((control) => {
    if (control.id === "student-search") return;
    control.disabled = busy;
  });
  if (!busy && state.bootstrap && !state.bootstrap.paste_supported) {
    document.querySelectorAll(".action-button.paste").forEach((button) => {
      button.disabled = true;
      button.title = "Desktop paste automation is currently available on Windows.";
    });
  }
}

function toast(message, error = false) {
  const item = document.createElement("div");
  item.className = `toast${error ? " is-error" : ""}`;
  item.textContent = message;
  elements.toastRegion.appendChild(item);
  window.setTimeout(() => item.remove(), error ? 7000 : 3500);
}

function setLog(text) {
  elements.activityLog.textContent = text || "Ready.";
  elements.activityLog.scrollTop = elements.activityLog.scrollHeight;
}

function renderTabs() {
  elements.sheetTabs.replaceChildren();
  state.bootstrap.sheets.forEach((sheetName) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `sheet-tab${sheetName === state.sheetName ? " is-active" : ""}`;
    button.textContent = sheetName;
    button.addEventListener("click", () => switchSheet(sheetName));
    elements.sheetTabs.appendChild(button);
  });
}

function visibleColumns() {
  if (state.columnView === "all") return state.columns;
  const identity = new Set(["First Name", "Last Name", "uid"]);
  return state.columns.filter(
    (column) => identity.has(column.key) || column.group === state.columnView
  );
}

function filteredRows() {
  if (state.searchScope === "all") return state.rows;
  const query = state.search.trim().toLowerCase();
  if (!query) return state.rows;
  return state.rows.filter((row) => {
    const values = row.values || {};
    const firstName = displayValue(values["First Name"]);
    const lastName = displayValue(values["Last Name"]);
    const searchable = [
      firstName,
      lastName,
      `${firstName} ${lastName}`.trim(),
      displayValue(values.uid),
    ];
    return searchable.some((value) => value.toLowerCase().includes(query));
  });
}

function closeSearchResults() {
  state.searchResults = [];
  elements.searchResults.replaceChildren();
  elements.searchResults.hidden = true;
}

function renderSearchResults() {
  elements.searchResults.replaceChildren();
  if (state.searchScope !== "all" || !state.search.trim()) {
    elements.searchResults.hidden = true;
    return;
  }

  if (!state.searchResults.length) {
    const empty = document.createElement("div");
    empty.className = "search-results-empty";
    empty.textContent = "No matching students in any class.";
    elements.searchResults.appendChild(empty);
  } else {
    state.searchResults.forEach((result) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "search-result";
      button.setAttribute("role", "option");

      const name = document.createElement("span");
      name.className = "search-result-name";
      name.textContent = result.full_name || "Unnamed student";

      const sheet = document.createElement("span");
      sheet.className = "search-result-sheet";
      sheet.textContent = result.sheet;

      const meta = document.createElement("span");
      meta.className = "search-result-meta";
      meta.textContent = result.uid ? `UID ${result.uid}` : "No UID";

      button.append(name, sheet, meta);
      button.addEventListener("click", () => openSearchResult(result));
      elements.searchResults.appendChild(button);
    });
  }
  elements.searchResults.hidden = false;
}

async function runAllClassSearch() {
  const query = state.search.trim();
  if (state.searchScope !== "all" || !query) {
    closeSearchResults();
    return;
  }

  const requestNumber = ++state.searchRequest;
  try {
    const payload = await api(`/api/search?q=${encodeURIComponent(query)}`);
    if (
      requestNumber !== state.searchRequest ||
      state.searchScope !== "all" ||
      query !== state.search.trim()
    ) {
      return;
    }
    state.searchResults = payload.results || [];
    renderSearchResults();
  } catch (error) {
    if (requestNumber !== state.searchRequest) return;
    closeSearchResults();
    toast(error.message, true);
  }
}

function scheduleStudentSearch() {
  window.clearTimeout(state.searchTimer);
  state.searchRequest += 1;
  renderTable();
  if (state.searchScope !== "all" || !state.search.trim()) {
    closeSearchResults();
    return;
  }
  state.searchTimer = window.setTimeout(runAllClassSearch, 180);
}

function scrollToStudent(studentId) {
  window.setTimeout(() => {
    const target = Array.from(elements.studentBody.querySelectorAll("tr")).find(
      (row) => row.dataset.studentId === studentId
    );
    target?.scrollIntoView({ block: "center" });
  }, 0);
}

async function openSearchResult(result) {
  if (state.busy) return;
  try {
    setBusy(true, "Opening student", `${result.full_name} in ${result.sheet}`);
    if (state.dirty || state.announcementDirty) await saveAll({ quiet: true });
    if (result.sheet !== state.sheetName) await loadSheet(result.sheet);

    state.searchScope = "current";
    elements.searchScope.value = "current";
    state.search = result.uid || result.full_name;
    elements.studentSearch.value = state.search;
    closeSearchResults();
    state.selected = new Set([`student:${result.student_id}`]);
    renderTable();
    scrollToStudent(result.student_id);
    toast(`Opened ${result.full_name} in ${result.sheet}.`);
  } catch (error) {
    toast(error.message, true);
  } finally {
    setBusy(false);
  }
}
function createSelectionCell(row) {
  const cell = document.createElement("td");
  cell.className = "selection-cell";
  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  checkbox.checked = state.selected.has(rowKey(row));
  checkbox.setAttribute("aria-label", `Select ${displayValue(row.values["First Name"])}`);
  checkbox.addEventListener("change", () => {
    if (checkbox.checked) state.selected.add(rowKey(row));
    else state.selected.delete(rowKey(row));
    renderSelection();
    checkbox.closest("tr").classList.toggle("is-selected", checkbox.checked);
  });
  cell.appendChild(checkbox);
  return cell;
}

function createEditor(row, column) {
  const current = displayValue(row.values[column.key]);
  let editor;
  if (column.kind === "select") {
    editor = document.createElement("select");
    const choices = [...column.options];
    if (current && !choices.includes(current)) choices.push(current);
    choices.forEach((choice) => {
      const option = document.createElement("option");
      option.value = choice;
      option.textContent = choice || " ";
      editor.appendChild(option);
    });
    editor.value = current;
  } else {
    editor = document.createElement("input");
    editor.type = "text";
    editor.value = current;
    if (column.kind === "long_text") {
      editor.classList.add("is-long");
      editor.title = current;
    }
  }

  editor.classList.add("cell-input");
  editor.setAttribute("aria-label", `${column.label} for ${displayValue(row.values["First Name"])}`);
  editor.addEventListener("input", () => {
    row.values[column.key] = editor.value;
    if (column.kind === "long_text") editor.title = editor.value;
    markDirty("sheet");
  });
  editor.addEventListener("change", () => {
    row.values[column.key] = editor.value;
    markDirty("sheet");
  });
  return editor;
}

function renderTable() {
  const columns = visibleColumns();
  const rows = filteredRows();
  const rosterNumberByKey = new Map(
    state.rows.map((row, index) => [rowKey(row), index + 1])
  );
  const headerRow = document.createElement("tr");

  const selectHeader = document.createElement("th");
  selectHeader.className = "selection-cell";
  const selectAll = document.createElement("input");
  selectAll.type = "checkbox";
  selectAll.setAttribute("aria-label", "Select all visible students");
  const visibleKeys = rows.map(rowKey);
  selectAll.checked = visibleKeys.length > 0 && visibleKeys.every((key) => state.selected.has(key));
  selectAll.indeterminate =
    visibleKeys.some((key) => state.selected.has(key)) && !selectAll.checked;
  selectAll.addEventListener("change", () => {
    visibleKeys.forEach((key) => {
      if (selectAll.checked) state.selected.add(key);
      else state.selected.delete(key);
    });
    renderTable();
  });
  selectHeader.appendChild(selectAll);
  headerRow.appendChild(selectHeader);

  const rowHeader = document.createElement("th");
  rowHeader.className = "row-number";
  rowHeader.textContent = "No.";
  rowHeader.title = "Student number";
  headerRow.appendChild(rowHeader);

  columns.forEach((column) => {
    const th = document.createElement("th");
    th.textContent = column.label;
    th.title = column.label;
    th.style.width = `${column.width}px`;
    th.style.minWidth = `${column.width}px`;
    headerRow.appendChild(th);
  });
  elements.studentHead.replaceChildren(headerRow);

  const fragment = document.createDocumentFragment();
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    if (row.student_id) tr.dataset.studentId = row.student_id;
    tr.classList.toggle("is-selected", state.selected.has(rowKey(row)));
    tr.appendChild(createSelectionCell(row));

    const rowNumber = document.createElement("td");
    rowNumber.className = "row-number";
    rowNumber.textContent = rosterNumberByKey.get(rowKey(row));
    rowNumber.title = row.excel_row
      ? `Automation row ${row.excel_row}`
      : "New student, not saved yet";
    rowNumber.classList.toggle("new-row-marker", !row.excel_row);
    tr.appendChild(rowNumber);

    columns.forEach((column) => {
      const td = document.createElement("td");
      td.style.width = `${column.width}px`;
      td.style.minWidth = `${column.width}px`;
      td.appendChild(createEditor(row, column));
      tr.appendChild(td);
    });
    fragment.appendChild(tr);
  });
  elements.studentBody.replaceChildren(fragment);
  elements.emptyState.hidden = rows.length > 0;
  elements.studentTable.hidden = rows.length === 0;
  renderSelection();
}

function renderSelection() {
  const selectedRows = state.rows.filter((row) => state.selected.has(rowKey(row))).length;
  const total = state.rows.length;
  elements.selectionCount.textContent = `${selectedRows} selected`;
  elements.rowCount.textContent = `${total} student${total === 1 ? "" : "s"}`;
}

function renderQuizTarget() {
  elements.quizGenerateTarget.textContent = `Writes to Quiz${state.quizNumber} Feedback`;
  elements.quizPasteTarget.textContent = `Uses Quiz${state.quizNumber} Feedback`;
  document.querySelectorAll(".segment").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.quiz === state.quizNumber);
  });
}

function applySheetData(data, preserve = null) {
  state.sheetName = data.sheet;
  state.columns = data.columns;
  state.rows = data.rows.map((row) => ({
    ...row,
    client_id: row.student_id ? null : crypto.randomUUID(),
  }));

  if (preserve) {
    state.selected.clear();
    state.rows.forEach((row) => {
      if (
        preserve.ids?.has(row.student_id) ||
        preserve.rows.has(Number(row.excel_row)) ||
        preserve.fingerprints.has(fingerprint(row))
      ) {
        state.selected.add(rowKey(row));
      }
    });
  } else {
    state.selected = new Set(state.rows.map(rowKey));
  }

  elements.announcementText.value = data.announcement || "";
  elements.announcementPath.textContent = data.announcement_path || "";
  elements.announcementPath.title = data.announcement_path || "";
  state.dirty = false;
  state.announcementDirty = false;
  setSaveState("Saved", "saved");
  renderTabs();
  renderTable();
}

function preservedSelection() {
  const ids = new Set();
  const rows = new Set();
  const fingerprints = new Set();
  state.rows.forEach((row) => {
    if (!state.selected.has(rowKey(row))) return;
    if (row.student_id) ids.add(row.student_id);
    if (row.excel_row) rows.add(Number(row.excel_row));
    fingerprints.add(fingerprint(row));
  });
  return { ids, rows, fingerprints };
}

async function loadSheet(sheetName, preserve = null) {
  const payload = await api(`/api/sheet?name=${encodeURIComponent(sheetName)}`);
  applySheetData(payload.data, preserve);
}

async function switchSheet(sheetName) {
  if (state.busy || sheetName === state.sheetName) return;
  try {
    if (state.dirty || state.announcementDirty) await saveAll({ quiet: true });
    setBusy(true, "Opening class", sheetName);
    await loadSheet(sheetName);
  } catch (error) {
    toast(error.message, true);
  } finally {
    setBusy(false);
  }
}

async function saveAnnouncement({ quiet = false } = {}) {
  if (!state.announcementDirty && quiet) return;
  const payload = await api("/api/announcement/save", {
    method: "POST",
    body: JSON.stringify({
      sheet: state.sheetName,
      text: elements.announcementText.value,
    }),
  });
  state.announcementDirty = false;
  elements.announcementPath.textContent = payload.path;
  elements.announcementPath.title = payload.path;
  if (!quiet) toast("Announcement saved.");
}

async function saveSheet({ quiet = false } = {}) {
  if (!state.dirty && quiet) return;
  const preserve = preservedSelection();
  const payload = await api("/api/sheet/save", {
    method: "POST",
    body: JSON.stringify({
      sheet: state.sheetName,
      rows: state.rows.map((row) => ({
        student_id: row.student_id,
        values: row.values,
      })),
    }),
  });

  const announcement = elements.announcementText.value;
  const announcementPath = elements.announcementPath.textContent;
  applySheetData(payload.data, preserve);
  elements.announcementText.value = announcement;
  elements.announcementPath.textContent = announcementPath;
  state.announcementDirty = false;
  if (!quiet) toast("Student sheet saved.");
}

async function saveAll({ quiet = false } = {}) {
  setSaveState("Saving", "neutral");
  try {
    await saveAnnouncement({ quiet: true });
    await saveSheet({ quiet: true });
    state.dirty = false;
    state.announcementDirty = false;
    setSaveState("Saved", "saved");
    if (!quiet) toast("All changes saved.");
  } catch (error) {
    setSaveState("Save failed", "dirty");
    throw error;
  }
}

function addStudent() {
  state.columnView = "student";
  elements.columnView.value = "student";
  const values = Object.fromEntries(state.columns.map((column) => [column.key, ""]));
  const row = {
    student_id: null,
    excel_row: null,
    client_id: crypto.randomUUID(),
    values,
  };
  state.rows.push(row);
  state.selected.add(rowKey(row));
  markDirty("sheet");
  renderTable();
  window.setTimeout(() => {
    elements.tableWrap.scrollTop = elements.tableWrap.scrollHeight;
    const newRow = elements.studentBody.lastElementChild;
    const firstInput = newRow?.querySelector(".cell-input");
    firstInput?.focus();
  }, 0);
}

function selectedActionRows() {
  return state.rows
    .filter((row) => row.excel_row && state.selected.has(rowKey(row)))
    .map((row) => Number(row.excel_row));
}

async function runAction(action) {
  const isPaste = action.startsWith("paste-");
  if (isPaste) {
    const selectedCount = state.rows.filter((row) => state.selected.has(rowKey(row))).length;
    const confirmed = window.confirm(
      `Paste for ${selectedCount} selected student${selectedCount === 1 ? "" : "s"}?\n\nThe robot will switch between apps and paste text. It will not send messages.`
    );
    if (!confirmed) return;
  }

  try {
    setBusy(
      true,
      isPaste ? "Supervised paste running" : "Generating feedback",
      isPaste ? "Keep WeCom and WhatsApp available." : "Writing results into the database."
    );
    setSaveState("Working", "neutral");
    await saveAll({ quiet: true });
    const rows = selectedActionRows();
    if (!rows.length) throw new Error("Select at least one saved student row.");
    if (action === "paste-announcement" && !elements.announcementText.value.trim()) {
      throw new Error("The announcement file is empty.");
    }

    setLog(`Running ${action} for rows ${rows.join(", ")}...`);
    const result = await api("/api/action", {
      method: "POST",
      body: JSON.stringify({
        action,
        sheet: state.sheetName,
        rows,
        quiz_number: state.quizNumber,
      }),
    });
    setLog(result.output || result.label);
    toast(result.label);
    await loadSheet(state.sheetName, {
      ids: new Set(),
      rows: new Set(rows),
      fingerprints: new Set(),
    });
  } catch (error) {
    const output = error.payload?.output || error.message;
    setLog(output);
    toast(error.message, true);
    setSaveState("Needs attention", "dirty");
  } finally {
    setBusy(false);
  }
}

function bindEvents() {
  elements.columnView.addEventListener("change", () => {
    state.columnView = elements.columnView.value;
    renderTable();
  });
  elements.studentSearch.addEventListener("input", () => {
    state.search = elements.studentSearch.value;
    scheduleStudentSearch();
  });
  elements.searchScope.addEventListener("change", () => {
    state.searchScope = elements.searchScope.value;
    elements.studentSearch.placeholder =
      state.searchScope === "all" ? "Name or UID across classes" : "Name or UID";
    scheduleStudentSearch();
  });
  elements.selectVisible.addEventListener("click", () => {
    filteredRows().forEach((row) => state.selected.add(rowKey(row)));
    renderTable();
  });
  elements.clearSelection.addEventListener("click", () => {
    state.selected.clear();
    renderTable();
  });
  elements.addRow.addEventListener("click", addStudent);
  elements.saveSheet.addEventListener("click", async () => {
    try {
      setBusy(true, "Saving database", state.sheetName);
      await saveAll();
    } catch (error) {
      toast(error.message, true);
    } finally {
      setBusy(false);
    }
  });
  elements.exportExcel.addEventListener("click", async () => {
    try {
      setBusy(true, "Exporting Excel", "Creating a separate workbook copy.");
      await saveAll({ quiet: true });
      const result = await api("/api/export", {
        method: "POST",
        body: JSON.stringify({}),
      });
      setLog(`${result.message}\n${result.path}`);
      toast("Excel export created.");
    } catch (error) {
      setLog(error.message);
      toast(error.message, true);
    } finally {
      setBusy(false);
    }
  });
  elements.saveAnnouncement.addEventListener("click", async () => {
    try {
      setBusy(true, "Saving announcement", state.sheetName);
      await saveAnnouncement();
      if (!state.dirty) setSaveState("Saved", "saved");
    } catch (error) {
      toast(error.message, true);
    } finally {
      setBusy(false);
    }
  });
  elements.announcementText.addEventListener("input", () => markDirty("announcement"));
  document.querySelectorAll(".segment").forEach((button) => {
    button.addEventListener("click", () => {
      state.quizNumber = button.dataset.quiz;
      renderQuizTarget();
    });
  });
  document.querySelectorAll("[data-action]").forEach((button) => {
    button.addEventListener("click", () => runAction(button.dataset.action));
  });
  elements.clearLog.addEventListener("click", () => setLog("Ready."));
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeSearchResults();
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
      event.preventDefault();
      elements.saveSheet.click();
    }
  });
  window.addEventListener("beforeunload", (event) => {
    if (!state.dirty && !state.announcementDirty) return;
    event.preventDefault();
    event.returnValue = "";
  });
}

async function initialize() {
  [
    "databasePath",
    "saveState",
    "sheetTabs",
    "columnView",
    "studentSearch",
    "searchScope",
    "searchResults",
    "addRow",
    "saveSheet",
    "exportExcel",
    "selectionCount",
    "rowCount",
    "selectVisible",
    "clearSelection",
    "tableWrap",
    "studentTable",
    "studentHead",
    "studentBody",
    "emptyState",
    "saveAnnouncement",
    "announcementText",
    "announcementPath",
    "quizGenerateTarget",
    "quizPasteTarget",
    "activityLog",
    "clearLog",
    "busyOverlay",
    "busyTitle",
    "busyDetail",
    "toastRegion",
  ].forEach((name) => {
    const id = name.replace(/[A-Z]/g, (letter) => `-${letter.toLowerCase()}`);
    elements[name] = byId(id);
  });

  bindEvents();
  try {
    setBusy(true, "Opening database", "Loading class sheets.");
    const bootstrap = await api("/api/bootstrap");
    state.bootstrap = bootstrap;
    elements.databasePath.textContent = bootstrap.database;
    elements.databasePath.title = bootstrap.database;
    renderTabs();
    renderQuizTarget();
    if (!bootstrap.default_sheet) throw new Error("The database has no class sheets.");
    await loadSheet(bootstrap.default_sheet);
    if (!bootstrap.paste_supported) {
      setLog("Database editing and generation are ready. Desktop paste automation is disabled on this operating system.");
    }
  } catch (error) {
    setLog(error.message);
    toast(error.message, true);
    setSaveState("Load failed", "dirty");
  } finally {
    setBusy(false);
  }
}

window.addEventListener("DOMContentLoaded", initialize);
