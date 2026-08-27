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
  attachments: [],
  selectedAttachments: new Set(),
  bulkUnchecked: new Set(),
  checkChannel: "auto",
  busy: false,
};

const CHANNEL_HINTS = {
  auto: "Uses each student's own channel, and when they have none set, tries WeCom then WhatsApp — filling in Parent Language from whichever one matched.",
  wecom: "Searches WeCom only. Parent Language is left alone, since forcing a channel says nothing about which language a family uses.",
  whatsapp: "Searches WhatsApp only. Parent Language is left alone.",
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
  // The blanket re-enable above would undo the "nothing ticked" disabled state.
  if (!busy) renderBulkSummary();
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

function formatFileSize(bytes) {
  const size = Number(bytes) || 0;
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function selectedAttachmentIds() {
  return state.attachments
    .filter((attachment) => state.selectedAttachments.has(attachment.id))
    .map((attachment) => attachment.id);
}

function renderAttachments() {
  elements.attachmentList.replaceChildren();
  const availableIds = new Set(state.attachments.map((attachment) => attachment.id));
  state.selectedAttachments = new Set(
    Array.from(state.selectedAttachments).filter((id) => availableIds.has(id))
  );

  state.attachments.forEach((attachment) => {
    const row = document.createElement("div");
    row.className = "attachment-row";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.className = "attachment-select";
    checkbox.checked = state.selectedAttachments.has(attachment.id);
    checkbox.setAttribute("aria-label", `Include ${attachment.name}`);
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) state.selectedAttachments.add(attachment.id);
      else state.selectedAttachments.delete(attachment.id);
      renderAttachments();
    });

    const file = document.createElement("div");
    file.className = "attachment-file";
    const link = document.createElement("a");
    link.className = "attachment-name";
    link.href = `/api/attachment?id=${encodeURIComponent(attachment.id)}`;
    link.target = "_blank";
    link.rel = "noopener";
    link.textContent = attachment.name;
    link.title = attachment.name;
    const meta = document.createElement("span");
    meta.className = "attachment-meta";
    meta.textContent = `${attachment.kind} | ${formatFileSize(attachment.size)}`;
    file.append(link, meta);

    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "attachment-remove";
    remove.textContent = "Remove";
    remove.setAttribute("aria-label", `Remove ${attachment.name}`);
    remove.addEventListener("click", () => deleteAttachment(attachment));

    row.append(checkbox, file, remove);
    elements.attachmentList.appendChild(row);
  });

  const selected = selectedAttachmentIds().length;
  const total = state.attachments.length;
  elements.attachmentSummary.textContent = total
    ? `${selected} of ${total} selected`
    : "0 files";
  elements.attachmentEmpty.hidden = total > 0;
}

async function uploadAttachments(files) {
  const selectedFiles = Array.from(files || []);
  if (!selectedFiles.length) return;
  try {
    setBusy(true, "Uploading attachments", `${selectedFiles.length} file(s) for ${state.sheetName}`);
    for (const file of selectedFiles) {
      const query = new URLSearchParams({ sheet: state.sheetName, name: file.name });
      const response = await fetch(`/api/attachment/upload?${query}`, {
        method: "POST",
        headers: { "Content-Type": file.type || "application/octet-stream" },
        body: file,
      });
      let payload;
      try {
        payload = await response.json();
      } catch {
        payload = { ok: false, error: `The server returned HTTP ${response.status}.` };
      }
      if (!response.ok || payload.ok === false) {
        throw new Error(payload.error || `Could not upload ${file.name}.`);
      }
      state.attachments = payload.attachments || [];
      state.selectedAttachments.add(payload.attachment.id);
      renderAttachments();
    }
    toast(`${selectedFiles.length} attachment${selectedFiles.length === 1 ? "" : "s"} uploaded.`);
  } catch (error) {
    toast(error.message, true);
  } finally {
    elements.attachmentInput.value = "";
    setBusy(false);
  }
}

async function deleteAttachment(attachment) {
  if (!window.confirm(`Remove ${attachment.name} from this class?`)) return;
  try {
    setBusy(true, "Removing attachment", attachment.name);
    const payload = await api("/api/attachment/delete", {
      method: "POST",
      body: JSON.stringify({
        sheet: state.sheetName,
        attachment_id: attachment.id,
      }),
    });
    state.attachments = payload.attachments || [];
    state.selectedAttachments.delete(attachment.id);
    renderAttachments();
    toast("Attachment removed.");
  } catch (error) {
    toast(error.message, true);
  } finally {
    setBusy(false);
  }
}
async function deleteStudent(row) {
  const name =
    `${displayValue(row.values["First Name"])} ${displayValue(row.values["Last Name"])}`.trim() ||
    "this student";

  if (!row.student_id) {
    if (!window.confirm(`Remove ${name} from the roster? This row has not been saved yet.`)) return;
    state.rows = state.rows.filter((item) => rowKey(item) !== rowKey(row));
    state.selected.delete(rowKey(row));
    markDirty("sheet");
    renderTable();
    return;
  }

  if (!window.confirm(`Delete ${name} from ${state.sheetName}? This cannot be undone.`)) return;

  try {
    setBusy(true, "Deleting student", name);
    // Any unsaved edits to other rows must be saved first -- the delete goes straight
    // to the database and reloading afterward would otherwise silently drop them.
    if (state.dirty || state.announcementDirty) await saveAll({ quiet: true });
    const payload = await api("/api/student/delete", {
      method: "POST",
      body: JSON.stringify({ sheet: state.sheetName, student_id: row.student_id }),
    });
    applySheetData(payload.data);
    toast(`${name} deleted.`);
  } catch (error) {
    toast(error.message, true);
  } finally {
    setBusy(false);
  }
}

async function deleteClass() {
  const sheetName = state.sheetName;
  const studentCount = state.rows.length;
  const confirmed = window.confirm(
    `Delete "${sheetName}" and all ${studentCount} student${studentCount === 1 ? "" : "s"} in it, ` +
      "including its announcement and attachments? This cannot be undone."
  );
  if (!confirmed) return;

  try {
    setBusy(true, "Deleting class", sheetName);
    const payload = await api("/api/class/delete", {
      method: "POST",
      body: JSON.stringify({ sheet: sheetName }),
    });
    state.bootstrap.sheets = payload.sheets;
    state.bootstrap.sheet_groups = payload.sheet_groups;
    state.bootstrap.default_sheet = payload.default_sheet;
    toast(`"${sheetName}" deleted.`);
    if (payload.default_sheet) {
      await loadSheet(payload.default_sheet);
    } else {
      state.sheetName = "";
      state.rows = [];
      state.columns = [];
      state.selected.clear();
      renderTabs();
      renderTable();
    }
  } catch (error) {
    toast(error.message, true);
  } finally {
    setBusy(false);
  }
}

function sheetGroups() {
  return state.bootstrap.sheet_groups || [{ semester: "", sheets: state.bootstrap.sheets || [] }];
}

function activeGroupIndex(groups) {
  const index = groups.findIndex((group) => group.sheets.includes(state.sheetName));
  return index === -1 ? 0 : index;
}

function renderTabs() {
  const groups = sheetGroups();
  const groupIndex = activeGroupIndex(groups);
  const activeGroup = groups[groupIndex];

  elements.semesterSelect.hidden = groups.length <= 1;
  if (groups.length > 1) {
    elements.semesterSelect.replaceChildren(
      ...groups.map((group, index) => {
        const option = document.createElement("option");
        option.value = String(index);
        option.textContent = group.semester || "Other Classes";
        return option;
      })
    );
    elements.semesterSelect.value = String(groupIndex);
  }

  elements.deleteSemester.hidden = !(groups.length > 1 && activeGroup.is_semester);
  elements.deleteSemester.textContent = `Delete "${activeGroup.semester}"`;

  elements.sheetTabs.replaceChildren();
  activeGroup.sheets.forEach((sheetName) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `sheet-tab${sheetName === state.sheetName ? " is-active" : ""}`;
    button.textContent = sheetName;
    button.addEventListener("click", () => switchSheet(sheetName));
    elements.sheetTabs.appendChild(button);
  });

  renderBulkClasses();
}

async function switchSemester(groupIndex) {
  const groups = sheetGroups();
  const group = groups[Number(groupIndex)];
  if (!group || !group.sheets.length) return;
  if (group.sheets.includes(state.sheetName)) return;
  await switchSheet(group.sheets[0]);
}

function renderChannelChoice() {
  document.querySelectorAll(".channel-segment").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.channel === state.checkChannel);
  });
  if (elements.channelHint) elements.channelHint.textContent = CHANNEL_HINTS[state.checkChannel];
}

function bulkSelectedSheets() {
  return Array.from(elements.bulkClassList.querySelectorAll("input[type=checkbox]"))
    .filter((box) => box.checked)
    .map((box) => box.dataset.sheet);
}

function renderBulkSummary() {
  if (!state.bootstrap || !elements.bulkClassList) return;
  const groups = sheetGroups();
  const group = groups[activeGroupIndex(groups)] || { student_counts: {} };
  const counts = group.student_counts || {};
  const selected = bulkSelectedSheets();
  const students = selected.reduce((total, sheet) => total + (counts[sheet] || 0), 0);
  elements.bulkSummary.textContent = `${selected.length} class(es), ${students} student(s)`;
  elements.bulkCheck.disabled = state.busy || selected.length === 0;
  elements.bulkToggleAll.textContent = selected.length ? "None" : "All";
}

function renderBulkClasses() {
  const groups = sheetGroups();
  const group = groups[activeGroupIndex(groups)];
  elements.bulkScope.textContent = group ? group.semester || "Other Classes" : "No classes";
  elements.bulkClassList.replaceChildren();

  const sheets = group ? group.sheets : [];
  const counts = (group && group.student_counts) || {};
  if (!sheets.length) {
    const empty = document.createElement("p");
    empty.className = "bulk-empty";
    empty.textContent = "No classes in this semester.";
    elements.bulkClassList.appendChild(empty);
  }

  sheets.forEach((sheetName) => {
    const label = document.createElement("label");
    label.className = "bulk-class";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    // Classes start ticked; anything the user unticks stays unticked when this list
    // re-renders (which happens on every sheet load), rather than silently resetting.
    checkbox.checked = !state.bulkUnchecked.has(sheetName);
    checkbox.dataset.sheet = sheetName;
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) state.bulkUnchecked.delete(sheetName);
      else state.bulkUnchecked.add(sheetName);
      renderBulkSummary();
    });

    const name = document.createElement("span");
    name.className = "bulk-class-name";
    name.textContent = sheetName;
    name.title = sheetName;

    const count = document.createElement("span");
    count.className = "bulk-class-count";
    count.textContent = String(counts[sheetName] ?? 0);

    label.appendChild(checkbox);
    label.appendChild(name);
    label.appendChild(count);
    elements.bulkClassList.appendChild(label);
  });

  renderBulkSummary();
}

async function runBulkGroupChatCheck() {
  if (state.busy) return;
  const sheets = bulkSelectedSheets();
  if (!sheets.length) {
    toast("Tick at least one class to check.", true);
    return;
  }
  const groups = sheetGroups();
  const group = groups[activeGroupIndex(groups)] || { student_counts: {} };
  const counts = group.student_counts || {};
  const students = sheets.reduce((total, sheet) => total + (counts[sheet] || 0), 0);
  // Each student is a separate WeCom/WhatsApp search, so a full semester takes a while
  // and holds onto the desktop app the whole time. Say so before starting.
  const minutes = Math.max(1, Math.round((students * 3) / 60));
  const confirmed = window.confirm(
    `Check group chats for ${students} student(s) across ${sheets.length} class(es)?\n\n` +
      `This searches WeCom and WhatsApp once per student and will take roughly ${minutes} minute(s), ` +
      `using those apps the whole time. It only reads and never sends anything.`
  );
  if (!confirmed) return;

  try {
    setBusy(
      true,
      "Checking group chats",
      `${students} student(s) across ${sheets.length} class(es). Keep WeCom and WhatsApp available.`
    );
    setSaveState("Working", "neutral");
    await saveAll({ quiet: true });
    setLog(`Checking group chats for: ${sheets.join(", ")}...`);

    const result = await api("/api/action", {
      method: "POST",
      body: JSON.stringify({
        action: "check-group-chat-bulk",
        sheets,
        channel: state.checkChannel,
      }),
    });
    setLog(result.output || result.label);
    toast(result.label);
    await loadSheet(state.sheetName, preservedSelection());
  } catch (error) {
    const output = error.payload?.output || error.message;
    setLog(output);
    toast(error.message, true);
    setSaveState("Needs attention", "dirty");
  } finally {
    setBusy(false);
  }
}

function currentSemesterParam() {
  const groups = sheetGroups();
  const group = groups[activeGroupIndex(groups)];
  if (!group) return null;
  // Matches UNASSIGNED_SEMESTER in database_store.py: the legacy classes that have
  // no semester_id, shown as "Other Classes".
  return group.is_semester ? group.semester : "__unassigned__";
}

async function downloadExport(type, label) {
  if (state.busy) return;
  const semester = currentSemesterParam();
  try {
    setBusy(true, `Preparing ${label.toLowerCase()}`, semester ? `Semester: ${semester}` : "");
    await saveAll({ quiet: true });

    const query = new URLSearchParams({ type });
    if (semester) query.set("semester", semester);
    const response = await fetch(`/api/export/download?${query}`);
    if (!response.ok) {
      let message = `The server returned HTTP ${response.status}.`;
      try {
        const payload = await response.json();
        message = payload.error || message;
      } catch {
        /* A non-JSON error body leaves the generic status message in place. */
      }
      throw new Error(message);
    }

    const blob = await response.blob();
    const disposition = response.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename="([^"]+)"/);
    const filename = match ? match[1] : `${type}_export.xlsx`;

    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);

    setLog(`${label} downloaded as ${filename}.`);
    toast(`${label} downloaded.`);
  } catch (error) {
    setLog(error.message);
    toast(error.message, true);
  } finally {
    setBusy(false);
  }
}

async function deleteCurrentSemester() {
  if (state.busy) return;
  const groups = sheetGroups();
  const group = groups[activeGroupIndex(groups)];
  if (!group || !group.is_semester) return;
  const semesterName = group.semester;

  try {
    setBusy(true, "Checking semester", `Looking up what "${semesterName}" contains.`);
    const preview = await api("/api/semester/delete", {
      method: "POST",
      body: JSON.stringify({ semester: semesterName, confirm: false }),
    });
    const plan = preview.plan;
    const confirmed = window.confirm(
      `Delete semester "${semesterName}"? This permanently removes ${plan.classes.length} class(es) and ${plan.student_count} student(s): ${plan.classes.join(", ")}.\n\nA database backup is made automatically, but this cannot be undone from the app.`
    );
    if (!confirmed) {
      toast("Cancelled. Nothing was deleted.");
      return;
    }

    setBusy(true, "Deleting semester", `Removing "${semesterName}" and its classes.`);
    await api("/api/semester/delete", {
      method: "POST",
      body: JSON.stringify({ semester: semesterName, confirm: true }),
    });
    toast(`Deleted "${semesterName}".`);

    const bootstrap = await api("/api/bootstrap");
    state.bootstrap = bootstrap;
    const remainingGroups = sheetGroups();
    const fallbackSheet = remainingGroups[0]?.sheets[0] || bootstrap.default_sheet;
    if (fallbackSheet) {
      await loadSheet(fallbackSheet);
    } else {
      renderTabs();
    }
  } catch (error) {
    toast(error.message, true);
  } finally {
    setBusy(false);
  }
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

  const actionsHeader = document.createElement("th");
  actionsHeader.className = "row-actions";
  headerRow.appendChild(actionsHeader);
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

    const actionsCell = document.createElement("td");
    actionsCell.className = "row-actions";
    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.className = "icon-button remove-student-row";
    deleteButton.textContent = "×";
    const studentLabel = displayValue(row.values["First Name"]) || "this student";
    deleteButton.title = `Delete ${studentLabel}`;
    deleteButton.setAttribute("aria-label", `Delete ${studentLabel}`);
    deleteButton.addEventListener("click", () => deleteStudent(row));
    actionsCell.appendChild(deleteButton);
    tr.appendChild(actionsCell);

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
  state.attachments = data.attachments || [];
  state.selectedAttachments = new Set(state.attachments.map((attachment) => attachment.id));
  state.dirty = false;
  state.announcementDirty = false;
  setSaveState("Saved", "saved");
  renderTabs();
  renderTable();
  renderAttachments();
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
  const preservedAttachments = new Set(state.selectedAttachments);
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
  state.selectedAttachments = new Set(
    state.attachments
      .filter((attachment) => preservedAttachments.has(attachment.id))
      .map((attachment) => attachment.id)
  );
  renderAttachments();
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
  const isGroupChatCheck = action === "check-group-chat";
  const usesDesktopApp = isPaste || isGroupChatCheck;
  const selectedCount = state.rows.filter((row) => state.selected.has(rowKey(row))).length;
  const attachmentIds = action === "paste-announcement" ? selectedAttachmentIds() : [];
  if (isPaste && attachmentIds.length && selectedCount !== 1) {
    toast("Select exactly one student when staging attachments for review.", true);
    return;
  }
  if (isPaste) {
    const fileDetail = attachmentIds.length
      ? ` and ${attachmentIds.length} selected file${attachmentIds.length === 1 ? "" : "s"}`
      : "";
    const confirmed = window.confirm(
      `Paste text${fileDetail} for ${selectedCount} selected student${selectedCount === 1 ? "" : "s"}?\n\nThe robot will switch between apps and prepare the message. It will not press Send.`
    );
    if (!confirmed) return;
  }

  try {
    setBusy(
      true,
      isPaste ? "Supervised paste running" : isGroupChatCheck ? "Checking group chats" : "Generating feedback",
      usesDesktopApp ? "Keep WeCom and WhatsApp available." : "Writing results into the database."
    );
    setSaveState("Working", "neutral");
    await saveAll({ quiet: true });
    const rows = selectedActionRows();
    if (!rows.length) throw new Error("Select at least one saved student row.");
    if (action === "paste-announcement" && !elements.announcementText.value.trim() && !attachmentIds.length) {
      throw new Error("Add announcement text, select at least one attachment, or both.");
    }

    setLog(`Running ${action} for rows ${rows.join(", ")}...`);
    const result = await api("/api/action", {
      method: "POST",
      body: JSON.stringify({
        action,
        sheet: state.sheetName,
        rows,
        quiz_number: state.quizNumber,
        attachment_ids: attachmentIds,
        channel: state.checkChannel,
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
  elements.semesterSelect.addEventListener("change", () => {
    switchSemester(elements.semesterSelect.value);
  });
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
  elements.deleteClass.addEventListener("click", deleteClass);
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
  elements.exportExcel.addEventListener("click", () => downloadExport("full", "Excel export"));
  elements.exportReport.addEventListener("click", () => downloadExport("report", "Report export"));
  elements.deleteSemester.addEventListener("click", deleteCurrentSemester);
  elements.bulkCheck.addEventListener("click", runBulkGroupChatCheck);
  document.querySelectorAll(".channel-segment").forEach((button) => {
    button.addEventListener("click", () => {
      state.checkChannel = button.dataset.channel;
      renderChannelChoice();
    });
  });
  elements.bulkToggleAll.addEventListener("click", () => {
    const boxes = Array.from(elements.bulkClassList.querySelectorAll("input[type=checkbox]"));
    const turnOn = bulkSelectedSheets().length === 0;
    boxes.forEach((box) => {
      box.checked = turnOn;
      if (turnOn) state.bulkUnchecked.delete(box.dataset.sheet);
      else state.bulkUnchecked.add(box.dataset.sheet);
    });
    renderBulkSummary();
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
  elements.addAttachment.addEventListener("click", () => elements.attachmentInput.click());
  elements.attachmentInput.addEventListener("change", () => uploadAttachments(elements.attachmentInput.files));
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
    "semesterSelect",
    "deleteSemester",
    "bulkScope",
    "bulkClassList",
    "bulkToggleAll",
    "bulkCheck",
    "bulkSummary",
    "channelHint",
    "sheetTabs",
    "columnView",
    "studentSearch",
    "searchScope",
    "searchResults",
    "addRow",
    "saveSheet",
    "exportExcel",
    "exportReport",
    "deleteClass",
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
    "addAttachment",
    "attachmentInput",
    "attachmentSummary",
    "attachmentList",
    "attachmentEmpty",
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
    renderChannelChoice();
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
