"use strict";

(() => {
  const state = {
    plan: null,
  };

  const elements = {};

  function byId(id) {
    return document.getElementById(id);
  }

  function toast(message, error = false) {
    const region = byId("toast-region");
    if (!region) return;
    const item = document.createElement("div");
    item.className = `toast${error ? " is-error" : ""}`;
    item.textContent = message;
    region.appendChild(item);
    window.setTimeout(() => item.remove(), error ? 7000 : 3500);
  }

  function setBusy(busy, title = "Working", detail = "") {
    const overlay = byId("busy-overlay");
    if (!overlay) return;
    overlay.hidden = !busy;
    const titleEl = byId("busy-title");
    const detailEl = byId("busy-detail");
    if (titleEl) titleEl.textContent = title;
    if (detailEl) detailEl.textContent = detail;
    [
      elements.importPreview,
      elements.importCommit,
      elements.importFile,
      elements.importSemester,
      elements.importOverwrite,
    ].forEach((control) => {
      if (control) control.disabled = busy;
    });
  }

  async function uploadFile(path, file, params) {
    const query = new URLSearchParams(params).toString();
    const response = await fetch(`${path}?${query}`, {
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
      throw new Error(payload.error || "The import request could not be completed.");
    }
    return payload;
  }

  function removalCount(plan) {
    return (plan.classes || []).reduce(
      (total, classPlan) => total + (classPlan.students_to_remove || []).length,
      0
    );
  }

  function selectedClassIds() {
    return Array.from(elements.importClasses.querySelectorAll("input[type=checkbox]"))
      .filter((box) => box.checked)
      .map((box) => box.dataset.classId);
  }

  function renderPlan(plan) {
    state.plan = plan;
    elements.importPlan.hidden = false;
    elements.importEmpty.hidden = true;

    const semesterState = plan.semester_exists ? "existing semester" : "will be created";
    elements.importSemesterSummary.textContent = `${plan.semester} (${semesterState})`;

    elements.importClasses.replaceChildren();
    plan.classes.forEach((classPlan) => {
      const newCount = classPlan.students.filter((student) => !student.already_present).length;
      const alreadyCount = classPlan.students.length - newCount;
      const toRemove = classPlan.students_to_remove || [];

      const label = document.createElement("label");
      label.className = "import-class";

      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.checked = true;
      checkbox.dataset.classId = classPlan.external_class_id;

      const info = document.createElement("div");
      const title = document.createElement("strong");
      title.textContent = `${classPlan.class_exists ? "" : "[NEW] "}${classPlan.name}`;
      const detail = document.createElement("span");
      detail.textContent = `${classPlan.weekly_time || "no weekly time listed"} · ${newCount} new student(s), ${alreadyCount} already on roster`;
      info.appendChild(title);
      info.appendChild(detail);

      label.appendChild(checkbox);
      label.appendChild(info);
      elements.importClasses.appendChild(label);

      if (toRemove.length) {
        const removedBlock = document.createElement("div");
        removedBlock.className = "import-class is-removed";
        const removedInfo = document.createElement("div");
        const removedTitle = document.createElement("strong");
        removedTitle.textContent = `${toRemove.length} student(s) will be PERMANENTLY REMOVED from ${classPlan.name}`;
        const removedDetail = document.createElement("span");
        removedDetail.textContent = toRemove
          .map((student) => `${student.first_name} ${student.last_name} (${student.uid})`)
          .join(", ");
        removedInfo.appendChild(removedTitle);
        removedInfo.appendChild(removedDetail);
        removedBlock.appendChild(removedInfo);
        elements.importClasses.appendChild(removedBlock);
      }
    });
  }

  async function preview() {
    const file = elements.importFile.files[0];
    const semester = elements.importSemester.value.trim();
    if (!file) {
      toast("Choose an enrollment .xlsx file first.", true);
      return;
    }
    if (!semester) {
      toast("Enter a semester name.", true);
      return;
    }

    try {
      setBusy(true, "Reading the file", "Matching classes and students against the roster.");
      const overwrite = elements.importOverwrite.checked;
      const result = await uploadFile("/api/import/preview", file, {
        semester,
        overwrite: overwrite ? "1" : "0",
      });
      renderPlan(result.plan);
      const removed = removalCount(result.plan);
      toast(
        removed
          ? `Preview ready. ${removed} student(s) would be removed — review before importing.`
          : "Preview ready. Review the classes below before importing."
      );
    } catch (error) {
      toast(error.message, true);
    } finally {
      setBusy(false);
    }
  }

  async function commit() {
    const file = elements.importFile.files[0];
    const semester = elements.importSemester.value.trim();
    if (!file || !semester || !state.plan) {
      toast("Preview the file before committing.", true);
      return;
    }
    const classIds = selectedClassIds();
    if (!classIds.length) {
      toast("Select at least one class to import.", true);
      return;
    }
    const overwrite = elements.importOverwrite.checked;
    const removed = removalCount(state.plan);
    const overwriteWarning = overwrite
      ? `\n\nOverwrite is on: any student in these classes not present in the file will be PERMANENTLY REMOVED${removed ? ` (${removed} shown in the preview)` : ""}.`
      : "";
    const confirmed = window.confirm(
      `Import ${classIds.length} class(es) into "${semester}"? This writes new classes and students into the database.${overwriteWarning}`
    );
    if (!confirmed) return;

    try {
      setBusy(true, "Importing", "Writing classes and students into the database.");
      let result = await uploadFile("/api/import/commit", file, {
        semester,
        class_ids: classIds.join(","),
        overwrite: overwrite ? "1" : "0",
      });

      if (result.requires_confirm_delete) {
        renderPlan(result.plan);
        const finalConfirm = window.confirm(
          `This will PERMANENTLY remove ${result.removal_count} student(s) not present in the uploaded file (see the highlighted list below). This cannot be undone from the app, though a database backup is made automatically. Continue?`
        );
        if (!finalConfirm) {
          toast("Import cancelled. Nothing was written.");
          return;
        }
        result = await uploadFile("/api/import/commit", file, {
          semester,
          class_ids: classIds.join(","),
          overwrite: "1",
          confirm_delete: "1",
        });
      }

      toast("Import committed. Reloading roster...");
      window.setTimeout(() => window.location.reload(), 800);
    } catch (error) {
      toast(error.message, true);
    } finally {
      setBusy(false);
    }
  }

  async function loadSemesters() {
    try {
      const response = await fetch("/api/semesters");
      const payload = await response.json();
      if (!payload.ok) return;
      elements.importSemesterOptions.replaceChildren();
      (payload.semesters || []).forEach((semester) => {
        const option = document.createElement("option");
        option.value = semester.name;
        elements.importSemesterOptions.appendChild(option);
      });
    } catch {
      /* Semester suggestions are a convenience; failing quietly is fine. */
    }
  }

  function init() {
    elements.importFile = byId("import-file");
    elements.importOverwrite = byId("import-overwrite");
    elements.importSemester = byId("import-semester");
    elements.importSemesterOptions = byId("import-semester-options");
    elements.importPreview = byId("import-preview");
    elements.importPlan = byId("import-plan");
    elements.importEmpty = byId("import-empty");
    elements.importSemesterSummary = byId("import-semester-summary");
    elements.importClasses = byId("import-classes");
    elements.importCommit = byId("import-commit");

    if (!elements.importPreview) return;

    elements.importPreview.addEventListener("click", preview);
    elements.importCommit.addEventListener("click", commit);
  }

  document.addEventListener("DOMContentLoaded", init);

  window.ImportStudentsView = {
    onShow: loadSemesters,
  };
})();
