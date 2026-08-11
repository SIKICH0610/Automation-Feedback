"use strict";

(() => {
  const state = {
    banks: [],
    bank: null,
    dirty: false,
    busy: false,
    loaded: false,
    view: "roster",
  };

  const elements = {};

  function byId(id) {
    return document.getElementById(id);
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
      throw new Error(payload.error || "The quiz bank action could not be completed.");
    }
    return payload;
  }

  function toast(message, error = false) {
    const item = document.createElement("div");
    item.className = `toast${error ? " is-error" : ""}`;
    item.textContent = message;
    elements.toastRegion.appendChild(item);
    window.setTimeout(() => item.remove(), error ? 7000 : 3500);
  }

  function setStatus(label, kind = "saved") {
    elements.bankSaveState.textContent = label;
    elements.bankSaveState.className = `bank-save-state is-${kind}`;
  }

  function setBusy(busy, title = "Working", detail = "Updating quiz bank.") {
    state.busy = busy;
    elements.busyOverlay.hidden = !busy;
    elements.busyTitle.textContent = title;
    elements.busyDetail.textContent = detail;
    elements.quizBankView.querySelectorAll("button, input, textarea").forEach((control) => {
      control.disabled = busy;
    });
  }

  function markDirty() {
    state.dirty = true;
    setStatus("Unsaved", "dirty");
  }

  function renderBankList() {
    elements.quizBankList.replaceChildren();
    state.banks.forEach((bank) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = `quiz-bank-button${bank.id === state.bank?.id ? " is-active" : ""}`;

      const name = document.createElement("span");
      name.className = "quiz-bank-button-name";
      name.textContent = bank.display_name;

      const count = document.createElement("span");
      count.className = "quiz-bank-button-count";
      count.textContent = `${bank.entry_count} question${bank.entry_count === 1 ? "" : "s"}`;

      button.append(name, count);
      button.addEventListener("click", () => loadBank(bank.id));
      elements.quizBankList.appendChild(button);
    });
  }

  function updateEntryCount() {
    const count = state.bank?.entries.length || 0;
    elements.bankEntryCount.textContent = `${count} question${count === 1 ? "" : "s"}`;
    elements.quizBankEmpty.hidden = count > 0;
  }

  function createTextField(labelText, value, onInput, options = {}) {
    const label = document.createElement("label");
    label.className = `field bank-field${options.className ? ` ${options.className}` : ""}`;

    const caption = document.createElement("span");
    caption.textContent = labelText;

    const control = options.multiline
      ? document.createElement("textarea")
      : document.createElement("input");
    if (options.multiline) {
      control.rows = options.rows || 5;
      control.value = value;
    } else {
      control.type = options.type || "text";
      control.value = value;
      if (options.min !== undefined) control.min = String(options.min);
    }
    if (options.placeholder) control.placeholder = options.placeholder;
    control.addEventListener("input", () => onInput(control.value));

    label.append(caption, control);
    return label;
  }

  function renderEntries() {
    elements.quizBankEntries.replaceChildren();
    if (!state.bank) {
      updateEntryCount();
      return;
    }

    state.bank.entries.forEach((entry) => {
      const article = document.createElement("article");
      article.className = "quiz-bank-entry";

      const header = document.createElement("div");
      header.className = "quiz-bank-entry-header";

      const questionField = createTextField(
        "Question",
        entry.question,
        (value) => {
          entry.question = value === "" ? "" : Number(value);
          markDirty();
        },
        { type: "number", min: 1, className: "question-field" }
      );

      const titleField = createTextField("Title", entry.title || "", (value) => {
        entry.title = value;
        markDirty();
      });
      titleField.classList.add("title-field");

      const remove = document.createElement("button");
      remove.type = "button";
      remove.className = "icon-button remove-bank-entry";
      remove.textContent = "×";
      remove.title = `Delete question ${entry.question}`;
      remove.setAttribute("aria-label", `Delete question ${entry.question}`);
      remove.addEventListener("click", () => {
        const confirmed = window.confirm(`Delete question ${entry.question} from this bank?`);
        if (!confirmed) return;
        state.bank.entries = state.bank.entries.filter((candidate) => candidate !== entry);
        markDirty();
        renderEntries();
      });

      header.append(questionField, titleField, remove);

      const fields = document.createElement("div");
      fields.className = "quiz-bank-entry-fields";
      fields.append(
        createTextField(
          "Matching keywords",
          (entry.patterns || []).join("\n"),
          (value) => {
            entry.patterns = value
              .split(/\r?\n/)
              .map((pattern) => pattern.trim())
              .filter(Boolean);
            markDirty();
          },
          { multiline: true, rows: 6, className: "patterns-field" }
        ),
        createTextField(
          "Chinese feedback",
          entry.chinese || "",
          (value) => {
            entry.chinese = value;
            markDirty();
          },
          { multiline: true, rows: 6 }
        ),
        createTextField(
          "English feedback",
          entry.english || "",
          (value) => {
            entry.english = value;
            markDirty();
          },
          { multiline: true, rows: 6 }
        )
      );

      article.append(header, fields);
      elements.quizBankEntries.appendChild(article);
    });
    updateEntryCount();
  }

  function applyBank(bank) {
    state.bank = {
      ...bank,
      entries: (bank.entries || []).map((entry) => ({
        ...entry,
        patterns: [...(entry.patterns || [])],
      })),
    };
    elements.quizBankTitle.textContent = state.bank.display_name;
    elements.bankDisplayName.value = state.bank.display_name;
    elements.bankClassName.value = state.bank.class_name;
    elements.bankQuizName.value = state.bank.quiz_name;
    state.dirty = false;
    setStatus("Saved", "saved");
    renderBankList();
    renderEntries();
    localStorage.setItem("feedbackQuizBankId", state.bank.id);
  }

  async function refreshBanks() {
    const payload = await api("/api/quiz-banks");
    state.banks = payload.banks || [];
    renderBankList();
  }

  async function loadBank(bankId) {
    if (state.busy || bankId === state.bank?.id) return;
    try {
      if (state.dirty) await saveBank({ quiet: true });
      setBusy(true, "Opening quiz bank", "Loading feedback entries.");
      const payload = await api(`/api/quiz-bank?id=${encodeURIComponent(bankId)}`);
      applyBank(payload.bank);
    } catch (error) {
      toast(error.message, true);
    } finally {
      setBusy(false);
    }
  }

  async function loadInitialBank() {
    await refreshBanks();
    if (!state.banks.length) {
      state.loaded = true;
      renderEntries();
      return;
    }
    const remembered = localStorage.getItem("feedbackQuizBankId");
    const defaultBank =
      state.banks.find((bank) => bank.id === remembered) ||
      state.banks.find((bank) => bank.quiz_name.toLowerCase() === "quiz 2") ||
      state.banks[0];
    const payload = await api(`/api/quiz-bank?id=${encodeURIComponent(defaultBank.id)}`);
    applyBank(payload.bank);
    state.loaded = true;
  }

  async function saveBank({ quiet = false } = {}) {
    if (!state.bank) return;
    if (!state.dirty && quiet) return;

    setStatus("Saving", "neutral");
    const payload = await api("/api/quiz-bank/save", {
      method: "POST",
      body: JSON.stringify({
        bank_id: state.bank.id,
        class_name: state.bank.class_name,
        quiz_name: state.bank.quiz_name,
        display_name: state.bank.display_name,
        entries: state.bank.entries,
      }),
    });
    state.banks = payload.banks || state.banks;
    applyBank(payload.bank);
    if (!quiet) toast("Quiz bank saved.");
  }

  function addEntry() {
    if (!state.bank) return;
    const used = new Set(
      state.bank.entries
        .map((entry) => Number(entry.question))
        .filter((question) => Number.isInteger(question) && question > 0)
    );
    let question = 1;
    while (used.has(question)) question += 1;
    state.bank.entries.push({
      entry_id: null,
      question,
      title: "",
      patterns: [],
      chinese: "",
      english: "",
    });
    markDirty();
    renderEntries();
    window.setTimeout(() => {
      const entries = elements.quizBankEntries.querySelectorAll(".quiz-bank-entry");
      const last = entries[entries.length - 1];
      last?.scrollIntoView({ block: "center" });
      last?.querySelector(".title-field input")?.focus();
    }, 0);
  }

  async function switchView(view) {
    if (state.busy || view === state.view) return;
    try {
      if (state.view === "quiz-banks" && state.dirty) {
        setBusy(true, "Saving quiz bank", "Saving before leaving the editor.");
        await saveBank({ quiet: true });
      }
      state.view = view;
      elements.rosterView.hidden = view !== "roster";
      elements.quizBankView.hidden = view !== "quiz-banks";
      if (elements.importView) elements.importView.hidden = view !== "import";
      document.querySelectorAll(".primary-tab").forEach((button) => {
        button.classList.toggle("is-active", button.dataset.view === view);
      });
      if (view === "quiz-banks" && !state.loaded) {
        setBusy(true, "Opening quiz banks", "Loading the feedback library.");
        await loadInitialBank();
      }
      if (view === "import") {
        window.ImportStudentsView?.onShow?.();
      }
    } catch (error) {
      toast(error.message, true);
    } finally {
      setBusy(false);
    }
  }

  function bindEvents() {
    document.querySelectorAll(".primary-tab").forEach((button) => {
      button.addEventListener("click", () => switchView(button.dataset.view));
    });
    elements.addBankEntry.addEventListener("click", addEntry);
    elements.saveQuizBank.addEventListener("click", async () => {
      try {
        setBusy(true, "Saving quiz bank", "Writing feedback entries to the database.");
        await saveBank();
      } catch (error) {
        setStatus("Save failed", "dirty");
        toast(error.message, true);
      } finally {
        setBusy(false);
      }
    });
    elements.bankDisplayName.addEventListener("input", () => {
      if (!state.bank) return;
      state.bank.display_name = elements.bankDisplayName.value;
      elements.quizBankTitle.textContent = state.bank.display_name || "Quiz Bank";
      markDirty();
    });
    elements.bankClassName.addEventListener("input", () => {
      if (!state.bank) return;
      state.bank.class_name = elements.bankClassName.value;
      markDirty();
    });
    elements.bankQuizName.addEventListener("input", () => {
      if (!state.bank) return;
      state.bank.quiz_name = elements.bankQuizName.value;
      markDirty();
    });

    document.addEventListener(
      "keydown",
      (event) => {
        if (
          state.view !== "quiz-banks" ||
          !(event.ctrlKey || event.metaKey) ||
          event.key.toLowerCase() !== "s"
        ) {
          return;
        }
        event.preventDefault();
        event.stopImmediatePropagation();
        elements.saveQuizBank.click();
      },
      true
    );
    window.addEventListener("beforeunload", (event) => {
      if (!state.dirty) return;
      event.preventDefault();
      event.returnValue = "";
    });
  }

  function initialize() {
    [
      "rosterView",
      "quizBankView",
      "importView",
      "quizBankList",
      "quizBankTitle",
      "bankEntryCount",
      "bankSaveState",
      "bankDisplayName",
      "bankClassName",
      "bankQuizName",
      "addBankEntry",
      "saveQuizBank",
      "quizBankEntries",
      "quizBankEmpty",
      "busyOverlay",
      "busyTitle",
      "busyDetail",
      "toastRegion",
    ].forEach((name) => {
      const id = name.replace(/[A-Z]/g, (letter) => `-${letter.toLowerCase()}`);
      elements[name] = byId(id);
    });
    bindEvents();
  }

  window.addEventListener("DOMContentLoaded", initialize);
})();
