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
  contentMode: "general",
  dirty: false,
  announcementDirty: false,
  recapDirty: false,
  recaps: { general: "", "1": "", "2": "" },
  quizRecapDirty: false,
  closingDirty: false,
  attachments: [],
  selectedAttachments: new Set(),
  bulkUnchecked: new Set(),
  checkChannel: "auto",
  busy: false,
  activeJobId: null,
};

const CHANNEL_HINTS = {
  auto: "WhatsApp 维护中：Auto 目前只检查 WeCom。恢复后将按学生自己的渠道路由。",
  wecom: "Searches WeCom only. Parent Language is left alone, since forcing a channel says nothing about which language a family uses.",
  whatsapp: "WhatsApp 功能维护中，暂不可用。",
};

// Temporary maintenance gate: WhatsApp checking is offline for now. Auto quietly
// downgrades to WeCom-only (with a notice); an explicit WhatsApp run is refused.
// Delete this function and its two call sites to restore WhatsApp.
function applyWhatsappMaintenance(channel) {
  if (channel === "whatsapp") {
    toast("WhatsApp 功能维护中，本次操作已取消。", true);
    return null;
  }
  if (channel === "auto") {
    toast("WhatsApp 维护中，本次仅检查 WeCom。");
    return "wecom";
  }
  return channel;
}

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
  else if (kind === "recap") state.recapDirty = true;
  else if (kind === "quizRecap") state.quizRecapDirty = true;
  else if (kind === "closing") state.closingDirty = true;
  else state.dirty = true;
  setSaveState("Unsaved changes", "dirty");
}

function setBusy(busy, title = "Working", detail = "Please keep this window open.") {
  state.busy = busy;
  elements.busyOverlay.hidden = !busy;
  elements.busyTitle.textContent = title;
  elements.busyDetail.textContent = detail;
  document.querySelectorAll("button, input, select, textarea").forEach((control) => {
    if (control.id === "student-search" || control.id === "busy-cancel") return;
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
  const lines = String(text || "Ready.").trim().split("\n");
  const last = lines[lines.length - 1] || "Ready.";
  elements.logLast.textContent = last.length > 46 ? `${last.slice(0, 46)}…` : last;
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
    if (state.dirty || state.announcementDirty || state.recapDirty || state.quizRecapDirty) await saveAll({ quiet: true });
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
  if (elements.channelSwitch) elements.channelSwitch.title = CHANNEL_HINTS[state.checkChannel];
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
  const nothing = state.busy || selected.length === 0;
  elements.bulkCheck.disabled = nothing;
  if (elements.bulkGenerate) elements.bulkGenerate.disabled = nothing;
  if (elements.bulkPaste) elements.bulkPaste.disabled = nothing;
  elements.bulkToggleAll.textContent = selected.length ? "None" : "All";
}

function renderBulkClasses() {
  const groups = sheetGroups();
  const group = groups[activeGroupIndex(groups)];
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

const BULK_ACTIONS = {
  "check-group-chat-bulk": {
    // Each student is a separate WeCom/WhatsApp search, so a full semester takes a
    // while and holds onto the desktop app the whole time. Say so before starting.
    confirm: (students, classes, minutes) =>
      `Check group chats for ${students} student(s) across ${classes} class(es)?\n\n` +
      `This searches WeCom and WhatsApp once per student and will take roughly ${minutes} minute(s), ` +
      `using those apps the whole time. It only reads and never sends anything.`,
    busyTitle: "Checking group chats",
    busyDetail: "Keep WeCom and WhatsApp available.",
    log: "Checking group chats for",
  },
  "generate-comments-bulk": {
    confirm: (students, classes) =>
      `Generate comments for ${students} student(s) across ${classes} class(es)?\n\n` +
      `This overwrites the Feedback column for every student on the ticked rosters, ` +
      `using each class's own lesson recap as the opening paragraph.`,
    busyTitle: "Generating comments",
    busyDetail: "Writing results into the database.",
    log: "Generating comments for",
  },
  "paste-comments-bulk": {
    confirm: (students, classes, minutes) =>
      `Paste comments for ${students} student(s) across ${classes} class(es)?\n\n` +
      `The robot will switch between apps and prepare each message, roughly ${minutes} minute(s) in total. ` +
      `It will not press Send.`,
    busyTitle: "Supervised paste running",
    busyDetail: "Keep WeCom and WhatsApp available.",
    log: "Pasting comments for",
  },
};

async function runBulkAction(action) {
  if (state.busy) return;
  const spec = BULK_ACTIONS[action];
  const sheets = bulkSelectedSheets();
  if (!sheets.length) {
    toast("Tick at least one class first.", true);
    return;
  }
  let channel = state.checkChannel;
  if (action === "check-group-chat-bulk") {
    channel = applyWhatsappMaintenance(channel);
    if (channel === null) return;
  }
  const groups = sheetGroups();
  const group = groups[activeGroupIndex(groups)] || { student_counts: {} };
  const counts = group.student_counts || {};
  const students = sheets.reduce((total, sheet) => total + (counts[sheet] || 0), 0);
  const minutes = Math.max(1, Math.round((students * 3) / 60));
  if (!window.confirm(spec.confirm(students, sheets.length, minutes))) return;

  try {
    setBusy(true, spec.busyTitle, `${students} student(s) across ${sheets.length} class(es). ${spec.busyDetail}`);
    setSaveState("Working", "neutral");
    await saveAll({ quiet: true });
    setLog(`${spec.log}: ${sheets.join(", ")}...`);

    const result = await runActionJob({
      action,
      sheets,
      channel,
    });
    setLog(result.output || result.label);
    toast(result.label, false);
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
    if (state.dirty || state.announcementDirty || state.recapDirty || state.quizRecapDirty) await saveAll({ quiet: true });
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
  } else if (column.kind === "long_text") {
    // Must be a textarea: an <input type=text> strips newlines on assignment, which
    // both displayed multi-paragraph feedback as one line and permanently destroyed
    // the paragraph breaks on the next save.
    editor = document.createElement("textarea");
    editor.rows = 1;
    editor.value = current;
    editor.classList.add("is-long");
    editor.title = current;
    // Class-driven rather than :focus-driven so the floating box also works when
    // the window itself is not the OS-focused window (e.g. a background webview).
    editor.addEventListener("focus", () => editor.classList.add("is-expanded"));
    editor.addEventListener("blur", () => editor.classList.remove("is-expanded"));
    attachCellAiButton(editor, column, row);
  } else {
    editor = document.createElement("input");
    editor.type = "text";
    editor.value = current;
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

// One recap box + one Generate/Paste pair, re-pointed by the content mode. The
// three recap texts stay separate records; the mode only decides which one the
// box is editing and which columns the actions target.
const MODE_META = {
  general: {
    title: "本节课内容回顾",
    tooltip: "生成评语的第一段。开头自动加“家长您好～”，你写的内容原样跟在后面；留空则用一句通用回顾。",
    column: "Feedback",
    view: "general",
    placeholder: "例如：今天的课程主要围绕三角形全等的判定、勾股定理的应用展开",
    generate: "generate-comments",
    paste: "paste-comments",
  },
  quiz1: {
    title: "Quiz 1 回顾",
    tooltip: "Quiz 1 反馈的第一段，只用于 Quiz 1 的消息。每次 quiz 各自保存一份回顾。",
    column: "Quiz1 Feedback",
    view: "quiz1",
    placeholder: "本次 quiz 的范围和整体情况，例如：满分 8 分，7 道选择 + 1 道证明，班级平均 6.9/8",
    generate: "generate-quiz-feedback",
    paste: "paste-quiz-feedback",
  },
  quiz2: {
    title: "Quiz 2 回顾",
    tooltip: "Quiz 2 反馈的第一段，只用于 Quiz 2 的消息。每次 quiz 各自保存一份回顾。",
    column: "Quiz2 Feedback",
    view: "quiz2",
    placeholder: "本次 quiz 的范围和整体情况，例如：4 道题 + 1 道 bonus，班级平均 6.4/9",
    generate: "generate-quiz-feedback",
    paste: "paste-quiz-feedback",
  },
};

function recapKey(mode = state.contentMode) {
  return mode === "general" ? "general" : mode === "quiz1" ? "1" : "2";
}

function modeQuizNumber() {
  return state.contentMode === "quiz2" ? "2" : "1";
}

function renderMode() {
  const meta = MODE_META[state.contentMode];
  document.querySelectorAll(".mode-switch .segment").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.mode === state.contentMode);
  });
  elements.recapTitle.textContent = meta.title;
  elements.recapTitle.title = meta.tooltip;
  elements.recapText.placeholder = meta.placeholder;
  elements.recapText.value = state.recaps[recapKey()] || "";
  elements.generateTarget.textContent = `Writes to ${meta.column}`;
  elements.closingBlock.hidden = state.contentMode !== "general";
  elements.pasteTarget.textContent = `Uses ${meta.column}`;
}

// Switching modes swaps which recap the one box edits: bank the on-screen text
// first, quietly saving it if it was dirty.
async function switchContentMode(mode) {
  if (mode === state.contentMode || !MODE_META[mode]) return;
  state.recaps[recapKey()] = elements.recapText.value;
  try {
    if (state.contentMode === "general" && state.recapDirty) {
      await saveLessonRecap({ quiet: true });
    } else if (state.contentMode !== "general" && state.quizRecapDirty) {
      await saveQuizRecap({ quiet: true });
    }
  } catch (error) {
    toast(error.message, true);
    return;
  }
  state.contentMode = mode;
  renderMode();
  // The table follows the mode, so what Generate writes is what the teacher sees.
  const view = MODE_META[mode].view;
  if (state.columnView !== view) {
    state.columnView = view;
    elements.columnView.value = view;
    renderTable();
  }
}

const AI_CELL_COLUMNS = new Set(["Remark for Student", "Additional Comment"]);

function aiAvailable() {
  return !!(state.bootstrap && state.bootstrap.ai && state.bootstrap.ai.available);
}

async function aiExpand(kind, sourceText) {
  const result = await api("/api/ai/expand", {
    method: "POST",
    body: JSON.stringify({ kind, text: sourceText }),
  });
  return result;
}

// Wires one heading AI button to a textarea: keywords in the box are replaced by
// the expanded draft, which the teacher then reviews and edits like any text.
function bindAiButton(button, textarea, kind, dirtyKind) {
  button.addEventListener("click", async () => {
    if (button.disabled) return;
    const source = textarea.value.trim();
    if (!source) {
      toast("先在框里写几个关键词。", true);
      return;
    }
    const originalLabel = button.textContent;
    button.disabled = true;
    button.textContent = "生成中…";
    try {
      const result = await aiExpand(kind, source);
      textarea.value = result.text;
      markDirty(dirtyKind);
      toast(`AI 草稿已生成（${result.elapsed}s），请过目修改后再保存。`);
    } catch (error) {
      toast(error.message, true);
    } finally {
      button.disabled = false;
      button.textContent = originalLabel;
    }
  });
}

// The long-text cell editor gets a floating AI button while expanded, for the
// keyword columns only (Remark for Student / Additional Comment).
function attachCellAiButton(editor, column, row) {
  if (!AI_CELL_COLUMNS.has(column.key)) return;
  let button = null;
  editor.addEventListener("focus", () => {
    if (!aiAvailable() || button) return;
    button = document.createElement("button");
    button.type = "button";
    button.className = "button secondary small ai-button cell-ai-button";
    button.textContent = "AI 润色";
    button.title = "关键词或速记润色成一小段家长反馈（本地 AI，生成后请过目修改）";
    // mousedown would blur the textarea and collapse the editor before click.
    button.addEventListener("mousedown", (event) => event.preventDefault());
    button.addEventListener("click", async () => {
      const source = editor.value.trim();
      if (!source) {
        toast("先在格子里写几个关键词。", true);
        return;
      }
      button.textContent = "生成中…";
      button.disabled = true;
      try {
        const result = await aiExpand("student", source);
        editor.value = result.text;
        row.values[column.key] = result.text;
        markDirty("sheet");
        toast(`AI 草稿已生成（${result.elapsed}s），请过目修改后再保存。`);
      } catch (error) {
        toast(error.message, true);
      } finally {
        button.textContent = "AI 润色";
        button.disabled = false;
        editor.focus();
      }
    });
    editor.parentElement.appendChild(button);
    positionCellAiButton(editor, button);
  });
  editor.addEventListener("blur", () => {
    // Give the button's own click a beat to land before removing it.
    window.setTimeout(() => {
      if (button && document.activeElement !== editor) {
        button.remove();
        button = null;
      }
    }, 250);
  });
}

function positionCellAiButton(editor, button) {
  // The expanded editor is absolutely positioned in its td; pin the button to
  // its bottom-right corner.
  button.style.right = "6px";
  button.style.top = "232px";
}

function switchRailTab(name) {
  document.querySelectorAll(".rail-tab").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.rail === name);
  });
  for (const panel of ["feedback", "announce", "bulk"]) {
    const node = byId(`rail-${panel}`);
    if (node) node.hidden = panel !== name;
  }
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
  state.recaps = {
    general: data.lesson_recap || "",
    "1": (data.quiz_recaps || {})["1"] || "",
    "2": (data.quiz_recaps || {})["2"] || "",
  };
  state.recapDirty = false;
  state.quizRecapDirty = false;
  elements.closingText.value = data.closing_note || "";
  state.closingDirty = false;
  renderMode();
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
    if (state.dirty || state.announcementDirty || state.recapDirty || state.quizRecapDirty) await saveAll({ quiet: true });
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

async function saveLessonRecap({ quiet = false } = {}) {
  if (!state.recapDirty && quiet) return;
  await api("/api/lesson-recap/save", {
    method: "POST",
    body: JSON.stringify({
      sheet: state.sheetName,
      text: elements.recapText.value,
    }),
  });
  state.recapDirty = false;
  if (!quiet) toast("Lesson recap saved.");
}

async function saveClosingNote({ quiet = false } = {}) {
  if (!state.closingDirty && quiet) return;
  await api("/api/closing-note/save", {
    method: "POST",
    body: JSON.stringify({ sheet: state.sheetName, text: elements.closingText.value }),
  });
  state.closingDirty = false;
  if (!quiet) toast("Closing paragraph saved.");
}

async function saveQuizRecap({ quiet = false } = {}) {
  if (!state.quizRecapDirty && quiet) return;
  const quizNumber = modeQuizNumber();
  const text = state.contentMode === "general" ? state.recaps[quizNumber] || "" : elements.recapText.value;
  await api("/api/quiz-recap/save", {
    method: "POST",
    body: JSON.stringify({ sheet: state.sheetName, quiz_number: quizNumber, text }),
  });
  state.recaps[quizNumber] = text;
  state.quizRecapDirty = false;
  if (!quiet) toast(`Quiz ${quizNumber} recap saved.`);
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
    await saveLessonRecap({ quiet: true });
    await saveQuizRecap({ quiet: true });
    await saveClosingNote({ quiet: true });
    await saveSheet({ quiet: true });
    state.dirty = false;
    state.announcementDirty = false;
    state.recapDirty = false;
    state.quizRecapDirty = false;
    state.closingDirty = false;
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

function formatEta(seconds) {
  if (!Number.isFinite(seconds) || seconds <= 0) return "";
  const minutes = Math.floor(seconds / 60);
  const rest = Math.round(seconds % 60);
  return minutes ? `${minutes}分${String(rest).padStart(2, "0")}秒` : `${rest}秒`;
}

function renderJobProgress(snapshot) {
  elements.busyProgress.hidden = false;
  const { done, total, detail, elapsed } = snapshot;
  const fill = elements.busyProgressFill;
  if (total > 0) {
    fill.classList.remove("is-indeterminate");
    fill.style.width = `${Math.min(100, Math.round((done / total) * 100))}%`;
    let text = `${done} / ${total}`;
    if (done >= 2 && done < total) {
      const eta = formatEta((elapsed / done) * (total - done));
      if (eta) text += ` · 预计还需 ${eta}`;
    }
    if (detail) text += ` · ${detail}`;
    elements.busyProgressText.textContent = text;
  } else {
    fill.classList.add("is-indeterminate");
    fill.style.width = "30%";
    elements.busyProgressText.textContent = detail || "";
  }
}

// Submits the action as a background job and polls it to completion, driving the
// progress bar and the cancel button. Resolves with the finished snapshot
// (state: finished | cancelled) and throws on state: failed.
async function runActionJob(payload) {
  const started = await api("/api/action", { method: "POST", body: JSON.stringify(payload) });
  state.activeJobId = started.job_id;
  elements.busyCancel.hidden = false;
  elements.busyCancel.disabled = false;
  elements.busyCancel.textContent = "取消 Cancel";
  try {
    for (;;) {
      await new Promise((resolve) => setTimeout(resolve, 700));
      const snapshot = await api(`/api/action/progress?id=${started.job_id}`);
      renderJobProgress(snapshot);
      if (snapshot.state === "failed") {
        const error = new Error(snapshot.error || "The action failed.");
        error.payload = snapshot;
        throw error;
      }
      if (snapshot.state !== "running") return snapshot;
    }
  } finally {
    state.activeJobId = null;
    elements.busyCancel.hidden = true;
    elements.busyProgress.hidden = true;
    elements.busyProgressFill.style.width = "0%";
  }
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
    const result = await runActionJob({
      action,
      sheet: state.sheetName,
      rows,
      quiz_number: modeQuizNumber(),
      attachment_ids: attachmentIds,
      channel: state.checkChannel,
    });
    setLog(result.output || result.label);
    toast(result.label, false);
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
  elements.busyCancel.addEventListener("click", async () => {
    if (!state.activeJobId) return;
    elements.busyCancel.disabled = true;
    elements.busyCancel.textContent = "正在取消…";
    try {
      await api("/api/action/cancel", {
        method: "POST",
        body: JSON.stringify({ id: state.activeJobId }),
      });
    } catch (error) {
      toast(error.message, true);
    }
  });
  elements.bulkCheck.addEventListener("click", () => runBulkAction("check-group-chat-bulk"));
  elements.bulkGenerate.addEventListener("click", () => runBulkAction("generate-comments-bulk"));
  elements.bulkPaste.addEventListener("click", () => runBulkAction("paste-comments-bulk"));
  document.querySelectorAll(".channel-segment").forEach((button) => {
    button.addEventListener("click", () => {
      state.checkChannel = button.dataset.channel;
      if (button.dataset.channel === "whatsapp") {
        toast("WhatsApp 功能维护中，该选项暂时无法运行。", true);
      }
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
  elements.saveRecap.addEventListener("click", async () => {
    try {
      setBusy(true, "Saving recap", state.sheetName);
      if (state.contentMode === "general") {
        await saveLessonRecap();
        if (state.closingDirty) await saveClosingNote({ quiet: true });
      } else {
        await saveQuizRecap();
      }
      if (!state.dirty) setSaveState("Saved", "saved");
    } catch (error) {
      toast(error.message, true);
    } finally {
      setBusy(false);
    }
  });
  elements.recapText.addEventListener("input", () =>
    markDirty(state.contentMode === "general" ? "recap" : "quizRecap")
  );
  elements.closingText.addEventListener("input", () => markDirty("closing"));
  bindAiButton(elements.aiRecap, elements.recapText, "recap", "recap");
  bindAiButton(elements.aiAnnounce, elements.announcementText, "announce", "announcement");
  elements.addAttachment.addEventListener("click", () => elements.attachmentInput.click());
  elements.attachmentInput.addEventListener("change", () => uploadAttachments(elements.attachmentInput.files));
  document.querySelectorAll(".mode-switch .segment").forEach((button) => {
    button.addEventListener("click", () => switchContentMode(button.dataset.mode));
  });
  document.querySelectorAll(".rail-tab").forEach((button) => {
    button.addEventListener("click", () => switchRailTab(button.dataset.rail));
  });
  elements.actionGenerate.addEventListener("click", () =>
    runAction(MODE_META[state.contentMode].generate)
  );
  elements.actionPaste.addEventListener("click", () =>
    runAction(MODE_META[state.contentMode].paste)
  );
  document.querySelectorAll("[data-action]").forEach((button) => {
    button.addEventListener("click", () => runAction(button.dataset.action));
  });
  elements.logToggle.addEventListener("click", () => {
    const open = elements.logBody.hidden;
    elements.logBody.hidden = !open;
    elements.logCaret.textContent = open ? "▾" : "▸";
  });
  elements.toolbarMore.addEventListener("click", (event) => {
    event.stopPropagation();
    elements.toolbarMenu.hidden = !elements.toolbarMenu.hidden;
  });
  document.addEventListener("click", (event) => {
    if (!elements.toolbarMenu.hidden && !elements.toolbarMenu.contains(event.target)) {
      elements.toolbarMenu.hidden = true;
    }
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
    if (!state.dirty && !state.announcementDirty && !state.recapDirty && !state.quizRecapDirty && !state.closingDirty) return;
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
    "bulkClassList",
    "bulkToggleAll",
    "bulkCheck",
    "bulkGenerate",
    "bulkPaste",
    "busyProgress",
    "busyProgressFill",
    "busyProgressText",
    "busyCancel",
    "bulkSummary",
    "channelSwitch",
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
    "saveRecap",
    "recapText",
    "addAttachment",
    "attachmentInput",
    "attachmentSummary",
    "attachmentList",
    "attachmentEmpty",
    "recapTitle",
    "closingBlock",
    "closingText",
    "aiRecap",
    "aiAnnounce",
    "generateTarget",
    "pasteTarget",
    "actionGenerate",
    "actionPaste",
    "logToggle",
    "logCaret",
    "logLast",
    "logBody",
    "toolbarMore",
    "toolbarMenu",
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
    renderMode();
    setLog("Ready.");
    if (aiAvailable()) {
      elements.aiRecap.hidden = false;
      elements.aiAnnounce.hidden = false;
    }
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
