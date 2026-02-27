/*
Manuscript Chunk Reviewer
Run:
- Open index.html directly in a browser, or
- From this folder: python -m http.server 5173
*/

(() => {
  const STORAGE_KEY = "mcr_state_v1";
  const CHUNK_SIZE_DEFAULT = 1000;
  const EMPTY_STATE_ILLUSTRATION = "assets/illustrations/empty-state.svg";
  const CONTENT = window.MCR_CONTENT || {};

  const $ = (id) => document.getElementById(id);

  const elements = {
    resumeBanner: $("resume-banner"),
    resumeMeta: $("resume-meta"),
    resumeBtn: $("resume-btn"),
    dismissBtn: $("dismiss-btn"),
    fileMeta: $("file-meta"),
    fileInput: $("file-input"),
    sessionInput: $("session-input"),
    loadSessionBtn: $("load-session-btn"),
    exportSessionBtn: $("export-session-btn"),
    exportTxtBtn: $("export-txt-btn"),
    exportDocxBtn: $("export-docx-btn"),
    docSettingsBtn: $("doc-settings-btn"),
    docSettingsModal: $("doc-settings-modal"),
    docSettingsClose: $("doc-settings-close"),
    chunkSize: $("chunk-size"),
    chunkSizeValue: $("chunk-size-value"),
    rechunkBtn: $("rechunk-btn"),
    layoutPreset: $("layout-preset"),
    engineSelect: $("engine-select"),
    modeSelect: $("mode-select"),
    applyScope: $("apply-scope"),
    analyzeBtn: $("analyze-btn"),
    suggestBtn: $("suggest-btn"),
    compareToggle: $("compare-toggle"),
    chunkInfo: $("chunk-info"),
    chunkStats: $("chunk-stats"),
    resetChunkBtn: $("reset-chunk-btn"),
    skipChunkBtn: $("skip-chunk-btn"),
    approveChunkBtn: $("approve-chunk-btn"),
    editorBody: $("editor-body"),
    stagingWrap: $("staging-wrap"),
    stagingBody: $("staging-body"),
    stagingPreview: $("staging-preview"),
    reviewPreviewToggle: $("review-preview-toggle"),
    historyBody: $("history-body"),
    docFontFamily: $("doc-font-family"),
    docFontSize: $("doc-font-size"),
    docLineSpacing: $("doc-line-spacing"),
    docMargins: $("doc-margins"),
    docIndentEnabled: $("doc-indent-enabled"),
    docIndentSize: $("doc-indent-size"),
    docSpaceBefore: $("doc-space-before"),
    docSpaceAfter: $("doc-space-after"),
  };

  let appState = createDefaultState();
  let pendingStoredState = null;
  let saveTimer = null;
  let activeParagraphId = null;

  const historyTimers = new Map();
  const dmp = window.diff_match_patch ? new diff_match_patch() : null;

  function getContent(path, fallback = "") {
    if (!path) return fallback;
    const parts = path.split(".");
    let current = CONTENT;
    for (const part of parts) {
      if (current && Object.prototype.hasOwnProperty.call(current, part)) {
        current = current[part];
      } else {
        return fallback;
      }
    }
    return current ?? fallback;
  }

  function applyContent() {
    document.querySelectorAll("[data-i18n]").forEach((el) => {
      const key = el.getAttribute("data-i18n");
      const value = getContent(key, el.textContent);
      if (value !== undefined && value !== null) {
        el.textContent = value;
      }
    });
  }

  function createDefaultState() {
    return {
      meta: {
        fileName: "",
        inputFormat: "",
        importedAt: "",
      },
      ui: {
        layoutPreset: "editor",
        chunkSizeTarget: CHUNK_SIZE_DEFAULT,
        engineType: "local",
        mode: "clean",
        compare: false,
        applyScope: "focused",
        reviewPreview: false,
      },
      docSettings: {
        fontFamily: "Times New Roman",
        fontSize: 12,
        lineSpacing: 1.2,
        marginsPreset: "clean",
        indentEnabled: true,
        firstLineIndentCm: 0.5,
        spacingBeforePt: 0,
        spacingAfterPt: 6,
      },
      sourceParagraphs: [],
      chunks: [],
      stagedParagraphs: [],
      currentChunkIndex: 0,
    };
  }

  function uid(prefix) {
    return `${prefix}_${Math.random().toString(36).slice(2, 8)}_${Date.now()}`;
  }

  function scheduleSave() {
    if (saveTimer) {
      clearTimeout(saveTimer);
    }
    saveTimer = setTimeout(() => {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(appState));
      } catch (err) {
        console.warn("Unable to save state", err);
      }
    }, 600);
  }

  function showResumeBanner() {
    if (!pendingStoredState) {
      elements.resumeBanner.classList.add("hidden");
      return;
    }
    const meta = pendingStoredState.meta || {};
    const fallback = getContent("messages.storedSession", "Stored session");
    const info = meta.fileName
      ? `${meta.fileName} | ${new Date(meta.importedAt).toLocaleString()}`
      : fallback;
    elements.resumeMeta.textContent = info;
    elements.resumeBanner.classList.remove("hidden");
  }

  function loadFromStorage() {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (!saved) {
      return;
    }
    try {
      pendingStoredState = JSON.parse(saved);
      showResumeBanner();
    } catch (err) {
      console.warn("Unable to parse stored state", err);
    }
  }

  function ensureHistory(para) {
    const fallback = typeof para.editedText === "string" ? para.editedText : para.originalText || "";
    if (!para.history || !Array.isArray(para.history.stack)) {
      para.history = { stack: [fallback], index: 0 };
      return;
    }
    if (!Array.isArray(para.history.stack) || !para.history.stack.length) {
      para.history.stack = [fallback];
      para.history.index = 0;
      return;
    }
    if (typeof para.history.index !== "number" || para.history.index < 0) {
      para.history.index = para.history.stack.length - 1;
    }
    const current = fallback;
    if (para.history.stack[para.history.index] !== current) {
      para.history.stack = [...para.history.stack.slice(0, para.history.index + 1), current];
      para.history.index = para.history.stack.length - 1;
    }
  }

  function hydrateState(state) {
    const defaults = createDefaultState();
    const merged = {
      ...defaults,
      ...state,
      meta: { ...defaults.meta, ...(state.meta || {}) },
      ui: { ...defaults.ui, ...(state.ui || {}) },
      docSettings: { ...defaults.docSettings, ...(state.docSettings || {}) },
      sourceParagraphs: Array.isArray(state.sourceParagraphs) ? state.sourceParagraphs : [],
      chunks: Array.isArray(state.chunks) ? state.chunks : [],
      stagedParagraphs: Array.isArray(state.stagedParagraphs) ? state.stagedParagraphs : [],
      currentChunkIndex: Number.isInteger(state.currentChunkIndex)
        ? state.currentChunkIndex
        : 0,
    };

    merged.chunks.forEach((chunk) => {
      if (!Array.isArray(chunk.paragraphs)) {
        chunk.paragraphs = [];
      }
      chunk.paragraphs.forEach((para) => {
        if (!para.id) {
          para.id = uid("p");
        }
        if (typeof para.originalText !== "string") {
          para.originalText = "";
        }
        if (typeof para.editedText !== "string") {
          para.editedText = para.originalText;
        }
        if (!Array.isArray(para.changeLog)) {
          para.changeLog = [];
        }
        if (!Array.isArray(para.flags)) {
          para.flags = [];
        }
        ensureHistory(para);
      });
    });

    return merged;
  }
  function setLayoutPreset(preset) {
    const map = {
      editor: ["70%", "30%"],
      balanced: ["50%", "50%"],
      review: ["40%", "60%"],
    };
    const [left, right] = map[preset] || map.editor;
    document.documentElement.style.setProperty("--split-left", left);
    document.documentElement.style.setProperty("--split-right", right);
  }

  function normalizeNewlines(text) {
    return text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  }

  function parseParagraphs(text) {
    const normalized = normalizeNewlines(text).trim();
    if (!normalized) {
      return [];
    }
    return normalized
      .split(/\n\s*\n+/)
      .map((para) => para.trim())
      .filter(Boolean);
  }

  function buildChunksFromParagraphs(paragraphs, targetSize) {
    const chunks = [];
    let current = [];
    let currentSize = 0;

    paragraphs.forEach((para) => {
      const len = para.length;
      if (current.length && currentSize + len > targetSize) {
        chunks.push(current);
        current = [para];
        currentSize = len;
      } else {
        current.push(para);
        currentSize += len;
      }
    });

    if (current.length) {
      chunks.push(current);
    }

    return chunks.map((chunkParas) => ({
      id: uid("chunk"),
      status: "pending",
      paragraphs: chunkParas.map((text, index) => ({
        id: uid(`p${index + 1}`),
        originalText: text,
        editedText: text,
        styleTag: "Normal",
        changeLog: [],
        flags: [],
        history: {
          stack: [text],
          index: 0,
        },
      })),
      changeLog: [],
      flags: [],
    }));
  }
  function recordHistory(para, value) {
    ensureHistory(para);
    const stack = para.history.stack;
    const current = stack[para.history.index];
    if (current === value) return;
    const trimmedStack = stack.slice(0, para.history.index + 1);
    trimmedStack.push(value);
    para.history.stack = trimmedStack;
    para.history.index = trimmedStack.length - 1;
  }

  function updateUndoRedoButtons(undoBtn, redoBtn, para) {
    if (undoBtn) undoBtn.disabled = !canUndo(para);
    if (redoBtn) redoBtn.disabled = !canRedo(para);
  }

  function scheduleHistory(para, value, undoBtn, redoBtn) {
    const existing = historyTimers.get(para.id);
    if (existing) {
      clearTimeout(existing);
    }
    const timer = setTimeout(() => {
      recordHistory(para, value);
      updateUndoRedoButtons(undoBtn, redoBtn, para);
      historyTimers.delete(para.id);
    }, 400);
    historyTimers.set(para.id, timer);
  }

  function canUndo(para) {
    ensureHistory(para);
    return para.history.index > 0;
  }

  function canRedo(para) {
    ensureHistory(para);
    return para.history.index < para.history.stack.length - 1;
  }

  function undoParagraph(para, textarea, undoBtn, redoBtn) {
    if (!canUndo(para)) return;
    para.history.index -= 1;
    para.editedText = para.history.stack[para.history.index];
    textarea.value = para.editedText;
    updateUndoRedoButtons(undoBtn, redoBtn, para);
    scheduleSave();
    renderCurrentChunkStatsOnly();
  }

  function redoParagraph(para, textarea, undoBtn, redoBtn) {
    if (!canRedo(para)) return;
    para.history.index += 1;
    para.editedText = para.history.stack[para.history.index];
    textarea.value = para.editedText;
    updateUndoRedoButtons(undoBtn, redoBtn, para);
    scheduleSave();
    renderCurrentChunkStatsOnly();
  }

  function isParagraphStaged(paragraphId) {
    return (appState.stagedParagraphs || []).some((item) => item.paragraphId === paragraphId);
  }

  function stageParagraph(chunk, para) {
    if (!chunk || !para) return;
    if (!appState.stagedParagraphs) {
      appState.stagedParagraphs = [];
    }
    if (isParagraphStaged(para.id)) return;
    appState.stagedParagraphs.push({
      id: uid("staged"),
      chunkId: chunk.id,
      paragraphId: para.id,
      createdAt: new Date().toISOString(),
    });
    scheduleSave();
    renderStaging();
  }

  function unstageParagraph(itemId) {
    appState.stagedParagraphs = (appState.stagedParagraphs || []).filter((item) => item.id !== itemId);
    scheduleSave();
    renderStaging();
    renderCurrentChunk();
  }

  function findParagraphLocation(paragraphId, chunkId) {
    const chunks = appState.chunks || [];
    if (chunkId) {
      const chunkIndex = chunks.findIndex((chunk) => chunk.id === chunkId);
      if (chunkIndex >= 0) {
        const chunk = chunks[chunkIndex];
        const paraIndex = chunk.paragraphs.findIndex((p) => p.id === paragraphId);
        if (paraIndex >= 0) {
          return { chunk, chunkIndex, paragraph: chunk.paragraphs[paraIndex], paraIndex };
        }
      }
    }
    for (let i = 0; i < chunks.length; i += 1) {
      const chunk = chunks[i];
      const paraIndex = chunk.paragraphs.findIndex((p) => p.id === paragraphId);
      if (paraIndex >= 0) {
        return { chunk, chunkIndex: i, paragraph: chunk.paragraphs[paraIndex], paraIndex };
      }
    }
    return null;
  }

  function jumpToParagraph(location) {
    if (!location) return;
    appState.currentChunkIndex = location.chunkIndex;
    activeParagraphId = location.paragraph.id;
    scheduleSave();
    render();
  }

  function importText(text, meta) {
    const paragraphs = parseParagraphs(text);
    appState = createDefaultState();
    appState.meta = { ...meta };
    appState.sourceParagraphs = paragraphs;
    appState.ui.chunkSizeTarget = appState.ui.chunkSizeTarget || CHUNK_SIZE_DEFAULT;
    appState.chunks = buildChunksFromParagraphs(paragraphs, appState.ui.chunkSizeTarget);
    appState.stagedParagraphs = [];
    appState.currentChunkIndex = 0;
    scheduleSave();
    render();
  }

  function handleFileImport(file) {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    const meta = {
      fileName: file.name,
      inputFormat: ext,
      importedAt: new Date().toISOString(),
    };

    if (ext === "docx") {
      if (!window.mammoth) {
        alert(getContent("alerts.missingMammoth", "Mammoth library is missing. Ensure mammoth.browser.min.js is loaded."));
        return;
      }
      const reader = new FileReader();
      reader.onload = (event) => {
        const arrayBuffer = event.target.result;
        mammoth
          .extractRawText({ arrayBuffer })
          .then((result) => {
            importText(result.value, meta);
          })
          .catch(() => {
            alert(getContent("alerts.docxParseError", "Unable to parse .docx file."));
          });
      };
      reader.readAsArrayBuffer(file);
      return;
    }

    const reader = new FileReader();
    reader.onload = (event) => {
      importText(event.target.result, meta);
    };
    reader.readAsText(file);
  }

  function currentChunk() {
    return appState.chunks[appState.currentChunkIndex] || null;
  }

  function chunkProgress() {
    const total = appState.chunks.length;
    const approved = appState.chunks.filter((c) => c.status === "approved").length;
    const skipped = appState.chunks.filter((c) => c.status === "skipped").length;
    const pending = total - approved - skipped;
    return { total, approved, skipped, pending };
  }

  function calcTextStats(text) {
    const trimmed = text.trim();
    const words = trimmed ? trimmed.split(/\s+/).length : 0;
    const chars = trimmed.length;
    return { words, chars };
  }

  function calcChunkStats(chunk) {
    if (!chunk) return { words: 0, chars: 0 };
    const text = chunk.paragraphs.map((p) => p.editedText).join("\n\n");
    return calcTextStats(text);
  }

  function renderFileMeta() {
    if (!appState.meta.fileName) {
      elements.fileMeta.textContent = getContent("brand.subtitleNoFile", "No file loaded");
      return;
    }
    const date = appState.meta.importedAt
      ? new Date(appState.meta.importedAt).toLocaleString()
      : "";
    elements.fileMeta.textContent = `${appState.meta.fileName} | ${appState.meta.inputFormat.toUpperCase()} | ${date}`;
  }

  function renderToolbar() {
    elements.chunkSize.value = appState.ui.chunkSizeTarget || CHUNK_SIZE_DEFAULT;
    elements.chunkSizeValue.textContent = elements.chunkSize.value;
    elements.layoutPreset.value = appState.ui.layoutPreset;
    elements.engineSelect.value = appState.ui.engineType;
    elements.modeSelect.value = appState.ui.mode;
    elements.applyScope.value = appState.ui.applyScope;
    elements.compareToggle.checked = appState.ui.compare;
    setLayoutPreset(appState.ui.layoutPreset);

    elements.docFontFamily.value = appState.docSettings.fontFamily;
    elements.docFontSize.value = appState.docSettings.fontSize;
    elements.docLineSpacing.value = appState.docSettings.lineSpacing;
    elements.docMargins.value = appState.docSettings.marginsPreset;
    elements.docIndentEnabled.checked = appState.docSettings.indentEnabled;
    elements.docIndentSize.value = appState.docSettings.firstLineIndentCm;
    elements.docSpaceBefore.value = appState.docSettings.spacingBeforePt;
    elements.docSpaceAfter.value = appState.docSettings.spacingAfterPt;
  }
  function renderCurrentChunk() {
    const chunk = currentChunk();
    if (!chunk) {
      elements.chunkInfo.textContent = getContent("messages.noChunks", "No chunks loaded");
      elements.chunkStats.textContent = "";
      const emptyText = getContent("messages.emptyState", "Import a file to begin chunk review.");
      elements.editorBody.innerHTML = `
        <div class="empty-state">
          <img src="${EMPTY_STATE_ILLUSTRATION}" alt="" aria-hidden="true" />
          <div>${emptyText}</div>
        </div>
      `;
      return;
    }

    const progress = chunkProgress();
    const chunkIndex = appState.currentChunkIndex + 1;
    const stats = calcChunkStats(chunk);
    const chunkLabel = getContent("labels.chunkLabel", "Chunk");

    elements.chunkInfo.textContent = `${chunkLabel} ${chunkIndex} of ${progress.total} | ${progress.approved} approved | ${progress.skipped} skipped`;
    elements.chunkStats.textContent = `${stats.words} words | ${stats.chars} chars`;

    elements.editorBody.innerHTML = "";

    chunk.paragraphs.forEach((para, index) => {
      ensureHistory(para);
      const block = document.createElement("div");
      block.className = "paragraph-block";
      if (para.id === activeParagraphId) {
        block.classList.add("active");
      }

      const header = document.createElement("div");
      header.className = "paragraph-header";

      const label = document.createElement("div");
      label.className = "paragraph-tag";
      const paraPrefix = getContent("labels.paragraphPrefix", "P");
      label.textContent = `${paraPrefix}${index + 1}`;

      const styleSelect = document.createElement("select");
      ["Normal", "Title", "Heading1", "Poem", "Quote"].forEach((tag) => {
        const option = document.createElement("option");
        option.value = tag;
        option.textContent = tag;
        if (para.styleTag === tag) {
          option.selected = true;
        }
        styleSelect.appendChild(option);
      });
      styleSelect.addEventListener("change", (event) => {
        para.styleTag = event.target.value;
        scheduleSave();
      });

      const tools = document.createElement("div");
      tools.className = "paragraph-tools";

      const sendBtn = document.createElement("button");
      sendBtn.type = "button";
      sendBtn.className = "btn icon-btn";
      sendBtn.textContent = getContent("buttons.send", "Send");
      sendBtn.disabled = isParagraphStaged(para.id);

      const undoBtn = document.createElement("button");
      undoBtn.type = "button";
      undoBtn.className = "btn icon-btn";
      undoBtn.textContent = getContent("buttons.undo", "Undo");

      const redoBtn = document.createElement("button");
      redoBtn.type = "button";
      redoBtn.className = "btn icon-btn";
      redoBtn.textContent = getContent("buttons.redo", "Redo");

      updateUndoRedoButtons(undoBtn, redoBtn, para);

      tools.appendChild(sendBtn);
      tools.appendChild(undoBtn);
      tools.appendChild(redoBtn);

      const right = document.createElement("div");
      right.className = "paragraph-controls";
      right.appendChild(styleSelect);
      right.appendChild(tools);

      header.appendChild(label);
      header.appendChild(right);

      const body = document.createElement("div");
      body.className = "paragraph-body" + (appState.ui.compare ? " compare" : "");

      if (appState.ui.compare) {
        const original = document.createElement("textarea");
        original.className = "readonly";
        original.readOnly = true;
        original.value = para.originalText;
        body.appendChild(original);
      }

      const edited = document.createElement("textarea");
      edited.value = para.editedText;
      edited.addEventListener("focus", () => {
        activeParagraphId = para.id;
      });
      edited.addEventListener("input", (event) => {
        para.editedText = event.target.value;
        scheduleHistory(para, para.editedText, undoBtn, redoBtn);
        scheduleSave();
        renderCurrentChunkStatsOnly();
      });
      edited.addEventListener("blur", () => {
        recordHistory(para, para.editedText);
        updateUndoRedoButtons(undoBtn, redoBtn, para);
      });
      edited.addEventListener("keydown", (event) => {
        const key = event.key.toLowerCase();
        if ((event.ctrlKey || event.metaKey) && !event.shiftKey && key === "z") {
          event.preventDefault();
          undoParagraph(para, edited, undoBtn, redoBtn);
        }
        if ((event.ctrlKey || event.metaKey) && (key === "y" || (event.shiftKey && key === "z"))) {
          event.preventDefault();
          redoParagraph(para, edited, undoBtn, redoBtn);
        }
      });
      body.appendChild(edited);

      sendBtn.addEventListener("click", () => {
        stageParagraph(chunk, para);
        sendBtn.disabled = true;
      });

      undoBtn.addEventListener("click", () => {
        undoParagraph(para, edited, undoBtn, redoBtn);
      });
      redoBtn.addEventListener("click", () => {
        redoParagraph(para, edited, undoBtn, redoBtn);
      });

      block.appendChild(header);
      block.appendChild(body);

      if (para.flags && para.flags.length) {
        const flags = document.createElement("div");
        flags.className = "flags-list";
        para.flags.forEach((flag) => {
          const item = document.createElement("div");
          item.className = `flag-item ${flag.severity || "info"}`;
          item.textContent = `${flag.message}`;
          flags.appendChild(item);
        });
        block.appendChild(flags);
      }

      elements.editorBody.appendChild(block);
    });
  }
  function renderCurrentChunkStatsOnly() {
    const chunk = currentChunk();
    if (!chunk) return;
    const stats = calcChunkStats(chunk);
    elements.chunkStats.textContent = `${stats.words} words | ${stats.chars} chars`;
  }

  function escapeHtml(text) {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function buildDiffHtml(before, after) {
    if (!dmp) {
      return `<div class="diff-block">${escapeHtml(after).replace(/\n/g, "<br>")}</div>`;
    }
    const diffs = dmp.diff_main(before, after);
    dmp.diff_cleanupSemantic(diffs);
    const html = diffs
      .map(([op, data]) => {
        const escaped = escapeHtml(data).replace(/\n/g, "<br>");
        if (op === 1) return `<span class="diff-add">${escaped}</span>`;
        if (op === -1) return `<span class="diff-del">${escaped}</span>`;
        return `<span class="diff-eq">${escaped}</span>`;
      })
      .join("");
    return `<div class="diff-block">${html}</div>`;
  }

  function aggregateChunkText(chunk, key) {
    return chunk.paragraphs.map((p) => p[key]).join("\n\n");
  }

  function aggregateChangeLog(chunk) {
    const log = [];
    chunk.paragraphs.forEach((para, index) => {
      para.changeLog.forEach((entry) => {
        log.push({ ...entry, paragraph: index + 1 });
      });
    });
    return log;
  }

  function aggregateFlags(chunk) {
    const flags = [];
    chunk.paragraphs.forEach((para, index) => {
      (para.flags || []).forEach((flag) => {
        flags.push({ ...flag, paragraph: index + 1 });
      });
    });
    return flags;
  }

  function renderStaging() {
    if (!elements.stagingBody) return;
    if (elements.reviewPreviewToggle) {
      elements.reviewPreviewToggle.checked = !!appState.ui.reviewPreview;
    }
    if (elements.stagingWrap) {
      elements.stagingWrap.classList.toggle("preview", !!appState.ui.reviewPreview);
    }
    elements.stagingBody.innerHTML = "";
    const staged = appState.stagedParagraphs || [];
    if (!staged.length) {
      const emptyText = getContent("messages.stagingEmpty", "No staged paragraphs yet.");
      elements.stagingBody.innerHTML = `<div class="empty-state">${emptyText}</div>`;
      renderStagingPreview(staged);
      return;
    }

    staged.forEach((item) => {
      const location = findParagraphLocation(item.paragraphId, item.chunkId);
      const card = document.createElement("div");
      card.className = "staging-card";

      const meta = document.createElement("div");
      meta.className = "staging-meta";

      const label = document.createElement("div");
      if (location) {
        const chunkLabel = getContent("labels.chunkLabel", "Chunk");
        const paraPrefix = getContent("labels.paragraphPrefix", "P");
        const styleTag = location.paragraph.styleTag || "Normal";
        label.textContent = `${chunkLabel} ${location.chunkIndex + 1} • ${paraPrefix}${location.paraIndex + 1} • ${styleTag}`;
      } else {
        label.textContent = getContent("messages.stagingMissing", "Source paragraph no longer available.");
      }

      const actions = document.createElement("div");
      actions.className = "staging-actions";

      if (location) {
        const jumpBtn = document.createElement("button");
        jumpBtn.type = "button";
        jumpBtn.className = "btn icon-btn";
        jumpBtn.textContent = getContent("buttons.jumpTo", "Go to");
        jumpBtn.addEventListener("click", () => {
          jumpToParagraph(location);
        });
        actions.appendChild(jumpBtn);
      }

      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.className = "btn icon-btn";
      removeBtn.textContent = getContent("buttons.remove", "Remove");
      removeBtn.addEventListener("click", () => {
        unstageParagraph(item.id);
      });
      actions.appendChild(removeBtn);

      meta.appendChild(label);
      meta.appendChild(actions);
      card.appendChild(meta);

      const text = document.createElement("div");
      text.className = "staging-text";
      text.textContent = location
        ? location.paragraph.editedText
        : getContent("messages.stagingMissing", "Source paragraph no longer available.");
      card.appendChild(text);

      elements.stagingBody.appendChild(card);
    });

    renderStagingPreview(staged);
  }

  function renderStagingPreview(staged) {
    if (!elements.stagingPreview) return;
    elements.stagingPreview.innerHTML = "";
    if (!staged.length) {
      const emptyText = getContent("messages.stagingEmpty", "No staged paragraphs yet.");
      elements.stagingPreview.innerHTML = `<div class="empty-state">${emptyText}</div>`;
      return;
    }

    const settings = appState.docSettings;
    const fontSizePx = Math.round(settings.fontSize * 1.333);
    const spacingBeforePx = Math.round(settings.spacingBeforePt * 1.333);
    const spacingAfterPx = Math.round(settings.spacingAfterPt * 1.333);
    const indentPx = settings.indentEnabled ? Math.round(settings.firstLineIndentCm * 37.8) : 0;

    const page = document.createElement("div");
    page.className = "book-page";
    page.style.fontFamily = settings.fontFamily;
    page.style.fontSize = `${fontSizePx}px`;
    page.style.lineHeight = settings.lineSpacing;

    const body = document.createElement("div");
    body.className = "book-page-body";

    staged.forEach((item) => {
      const location = findParagraphLocation(item.paragraphId, item.chunkId);
      const p = document.createElement("p");
      const styleTag = location?.paragraph?.styleTag || "Normal";
      const text = location?.paragraph?.editedText || getContent("messages.stagingMissing", "Source paragraph no longer available.");

      const tag = styleTag.toLowerCase();
      if (tag === "title") p.classList.add("title");
      if (tag === "heading1") p.classList.add("heading");
      if (tag === "quote") p.classList.add("quote");
      if (tag === "poem") p.classList.add("poem");

      p.classList.add("book-para");
      p.textContent = tag === "poem" ? text : text.replace(/\n+/g, " ");
      p.style.marginTop = `${spacingBeforePx}px`;
      p.style.marginBottom = `${spacingAfterPx}px`;

      if (indentPx && !["title", "heading1", "quote"].includes(tag)) {
        p.style.textIndent = `${indentPx}px`;
      }

      body.appendChild(p);
    });

    page.appendChild(body);
    elements.stagingPreview.appendChild(page);
  }

  function renderHistory() {
    const approved = appState.chunks.filter((c) => c.status === "approved");
    elements.historyBody.innerHTML = "";

    if (!approved.length) {
      const emptyText = getContent("messages.historyEmpty", "Approved chunks will appear here with diffs and trace logs.");
      elements.historyBody.innerHTML = `<div class="empty-state">${emptyText}</div>`;
      return;
    }

    approved.forEach((chunk, index) => {
      const details = document.createElement("details");
      details.open = index === approved.length - 1;

      const summary = document.createElement("summary");
      const stats = calcTextStats(aggregateChunkText(chunk, "editedText"));
      const chunkLabel = getContent("labels.chunkLabel", "Chunk");
      summary.textContent = `${chunkLabel} ${appState.chunks.indexOf(chunk) + 1} | ${chunk.paragraphs.length} paragraphs | ${stats.words} words`;
      details.appendChild(summary);

      const body = document.createElement("div");
      body.className = "accordion-body";

      const diffHtml = buildDiffHtml(
        aggregateChunkText(chunk, "originalText"),
        aggregateChunkText(chunk, "editedText")
      );
      const diffWrapper = document.createElement("div");
      diffWrapper.innerHTML = diffHtml;
      body.appendChild(diffWrapper);

      const log = chunk.changeLog.length ? chunk.changeLog : aggregateChangeLog(chunk);
      const logList = document.createElement("div");
      logList.className = "log-list";
      if (!log.length) {
        const empty = document.createElement("div");
        empty.className = "empty-state";
        empty.textContent = getContent("messages.noChangeLog", "No change log entries.");
        logList.appendChild(empty);
      } else {
        log.forEach((entry) => {
          const item = document.createElement("div");
          item.className = "log-entry";

          const title = document.createElement("div");
          title.className = "log-entry-title";
          const paraPrefix = getContent("labels.paragraphPrefix", "P");
          title.textContent = `${entry.summary}${entry.paragraph ? ` | ${paraPrefix}${entry.paragraph}` : ""}`;

          const reason = document.createElement("div");
          reason.className = "log-entry-reason";
          reason.textContent = entry.reason || "";

          const snippet = document.createElement("div");
          snippet.className = "log-snippet";
          snippet.textContent = `Before: ${entry.beforeSnippet || ""} | After: ${entry.afterSnippet || ""}`;

          item.appendChild(title);
          item.appendChild(reason);
          item.appendChild(snippet);
          logList.appendChild(item);
        });
      }
      body.appendChild(logList);

      const flags = chunk.flags.length ? chunk.flags : aggregateFlags(chunk);
      const flagList = document.createElement("div");
      flagList.className = "flags-list";
      if (!flags.length) {
        const empty = document.createElement("div");
        empty.className = "empty-state";
        empty.textContent = getContent("messages.noFlags", "No flags.");
        flagList.appendChild(empty);
      } else {
        flags.forEach((flag) => {
          const item = document.createElement("div");
          item.className = `flag-item ${flag.severity || "info"}`;
          const paraPrefix = getContent("labels.paragraphPrefix", "P");
          item.textContent = `${paraPrefix}${flag.paragraph}: ${flag.message}`;
          flagList.appendChild(item);
        });
      }
      body.appendChild(flagList);

      details.appendChild(body);
      elements.historyBody.appendChild(details);
    });
  }

  function render() {
    renderFileMeta();
    renderToolbar();
    renderCurrentChunk();
    renderStaging();
    renderHistory();
  }
  class LocalRulesEngine {
    analyzeParagraph(text) {
      const flags = [];
      const changeLog = [];
      const cleaned = text.trim();

      const sentenceParts = cleaned.split(/[.!?]/).filter(Boolean);
      const longest = sentenceParts.reduce((max, part) => Math.max(max, part.length), 0);
      if (longest > 200) {
        flags.push({
          type: "long_sentence",
          message: "Possible run-on sentence (>200 characters).",
          severity: "warn",
        });
        changeLog.push({
          type: "analysis",
          summary: "Possible run-on sentence",
          reason: "Detected a sentence segment longer than 200 characters.",
          beforeSnippet: snippet(cleaned),
          afterSnippet: "",
          scope: "paragraph",
        });
      }

      const connectorCount = (cleaned.match(/\b(and|y)\b/gi) || []).length;
      if (connectorCount >= 6) {
        flags.push({
          type: "connector_overuse",
          message: "Heavy connector chaining detected (and/y).",
          severity: "info",
        });
        changeLog.push({
          type: "analysis",
          summary: "Connector chaining",
          reason: "Multiple connector tokens detected in one paragraph.",
          beforeSnippet: snippet(cleaned),
          afterSnippet: "",
          scope: "paragraph",
        });
      }

      if (cleaned && !/[.!?\"]$/.test(cleaned)) {
        flags.push({
          type: "missing_punct",
          message: "Paragraph may be missing terminal punctuation.",
          severity: "info",
        });
        changeLog.push({
          type: "analysis",
          summary: "Possible missing punctuation",
          reason: "Paragraph does not end with punctuation.",
          beforeSnippet: snippet(cleaned),
          afterSnippet: "",
          scope: "paragraph",
        });
      }

      return { flags, changeLog };
    }

    suggestParagraph(text, mode) {
      let suggestedText = text;
      const changeLog = [];

      const before = suggestedText;
      const trimmed = suggestedText.trimEnd();
      if (trimmed && !/[.!?\"]$/.test(trimmed)) {
        suggestedText = trimmed + ".";
        changeLog.push({
          type: "suggestion",
          summary: "Added terminal punctuation",
          reason: "Paragraph lacked ending punctuation.",
          beforeSnippet: snippet(before),
          afterSnippet: snippet(suggestedText),
          scope: "paragraph",
        });
      }

      if (mode === "clean") {
        const splitResult = splitLongSentence(suggestedText);
        if (splitResult.changed) {
          suggestedText = splitResult.text;
          changeLog.push({
            type: "suggestion",
            summary: "Split a long sentence",
            reason: "A very long sentence was split conservatively at a comma.",
            beforeSnippet: snippet(splitResult.before),
            afterSnippet: snippet(splitResult.text),
            scope: "paragraph",
          });
        }
      }

      if (mode === "clean") {
        const connectorResult = softenConnectorOveruse(suggestedText);
        if (connectorResult.changed) {
          suggestedText = connectorResult.text;
          changeLog.push({
            type: "suggestion",
            summary: "Reduced connector chaining",
            reason: "Replaced one connector to reduce run-on flow.",
            beforeSnippet: snippet(connectorResult.before),
            afterSnippet: snippet(connectorResult.text),
            scope: "paragraph",
          });
        }
      }

      const spaced = suggestedText.replace(/\s{2,}/g, " ");
      if (spaced !== suggestedText && mode !== "translation") {
        const beforeSpace = suggestedText;
        suggestedText = spaced;
        changeLog.push({
          type: "suggestion",
          summary: "Normalized spacing",
          reason: "Collapsed repeated spaces.",
          beforeSnippet: snippet(beforeSpace),
          afterSnippet: snippet(spaced),
          scope: "paragraph",
        });
      }

      const analysis = this.analyzeParagraph(suggestedText);
      return {
        suggestedText,
        changeLog,
        flags: analysis.flags,
      };
    }
  }

  class AIEngine {
    analyzeParagraph(text) {
      return {
        flags: [{ type: "ai", message: "AI engine not configured.", severity: "info" }],
        changeLog: [
          {
            type: "analysis",
            summary: "AI engine placeholder",
            reason: "No API endpoint configured.",
            beforeSnippet: snippet(text),
            afterSnippet: "",
            scope: "paragraph",
          },
        ],
      };
    }

    suggestParagraph(text) {
      return {
        suggestedText: text,
        changeLog: [
          {
            type: "suggestion",
            summary: "AI suggestion skipped",
            reason: "AI engine is disabled for offline MVP.",
            beforeSnippet: snippet(text),
            afterSnippet: snippet(text),
            scope: "paragraph",
          },
        ],
        flags: [{ type: "ai", message: "AI engine not configured.", severity: "info" }],
      };
    }
  }

  const engines = {
    local: new LocalRulesEngine(),
    ai: new AIEngine(),
  };

  function snippet(text, limit = 80) {
    const cleaned = text.replace(/\s+/g, " ").trim();
    if (cleaned.length <= limit) return cleaned;
    const slicePoint = Math.max(0, limit - 3);
    return cleaned.slice(0, slicePoint) + "...";
  }

  function splitLongSentence(text) {
    const sentences = text.split(/([.!?])/);
    let combined = "";
    let changed = false;

    for (let i = 0; i < sentences.length; i += 2) {
      const segment = sentences[i] || "";
      const punctuation = sentences[i + 1] || "";
      if (!changed && segment.length > 240 && segment.includes(",")) {
        const midpoint = Math.floor(segment.length / 2);
        const commaIndex = findNearestComma(segment, midpoint);
        if (commaIndex > 0) {
          const before = segment.slice(0, commaIndex + 1);
          const after = segment.slice(commaIndex + 1).trimStart();
          combined += `${before} ${after ? after.charAt(0).toUpperCase() + after.slice(1) : ""}`;
          combined = combined.replace(/,\s*/, ". ");
          combined += punctuation;
          changed = true;
          continue;
        }
      }
      combined += segment + punctuation;
    }

    return { changed, text: combined || text, before: text };
  }

  function findNearestComma(segment, target) {
    const indices = [];
    for (let i = 0; i < segment.length; i += 1) {
      if (segment[i] === ",") indices.push(i);
    }
    if (!indices.length) return -1;
    return indices.reduce((closest, index) => {
      return Math.abs(index - target) < Math.abs(closest - target) ? index : closest;
    }, indices[0]);
  }

  function softenConnectorOveruse(text) {
    const matches = [...text.matchAll(/\b(and|y)\b/gi)];
    if (matches.length < 7) {
      return { changed: false, text, before: text };
    }
    const midpoint = Math.floor(matches.length / 2);
    const match = matches[midpoint];
    if (!match) {
      return { changed: false, text, before: text };
    }
    const start = match.index;
    const end = start + match[0].length;
    const before = text;
    const updated = `${text.slice(0, start)};${text.slice(end)}`;
    return { changed: true, text: updated, before };
  }

  function getScopeParagraphs() {
    const chunk = currentChunk();
    if (!chunk) return [];
    if (appState.ui.applyScope === "focused" && activeParagraphId) {
      const focused = chunk.paragraphs.find((p) => p.id === activeParagraphId);
      return focused ? [focused] : chunk.paragraphs;
    }
    return chunk.paragraphs;
  }

  function updateAnalysisEntries(paragraph, analysis) {
    paragraph.flags = analysis.flags;
    paragraph.changeLog = paragraph.changeLog.filter((entry) => entry.type !== "analysis");
    paragraph.changeLog.push(...analysis.changeLog);
  }

  function runAnalyze() {
    const engine = engines[appState.ui.engineType];
    if (!engine) return;
    const paragraphs = getScopeParagraphs();
    paragraphs.forEach((para) => {
      const result = engine.analyzeParagraph(para.editedText);
      updateAnalysisEntries(para, result);
    });
    scheduleSave();
    render();
  }

  function runSuggest() {
    const engine = engines[appState.ui.engineType];
    if (!engine) return;
    const paragraphs = getScopeParagraphs();
    paragraphs.forEach((para) => {
      const result = engine.suggestParagraph(para.editedText, appState.ui.mode);
      para.editedText = result.suggestedText;
      para.flags = result.flags;
      para.changeLog.push(...result.changeLog);
      recordHistory(para, para.editedText);
    });
    scheduleSave();
    render();
  }

  function resetCurrentChunk() {
    const chunk = currentChunk();
    if (!chunk) return;
    chunk.paragraphs.forEach((para) => {
      para.editedText = para.originalText;
      para.changeLog = [];
      para.flags = [];
      para.history = { stack: [para.originalText], index: 0 };
    });
    scheduleSave();
    render();
  }
  function skipCurrentChunk() {
    const chunk = currentChunk();
    if (!chunk) return;
    chunk.status = "skipped";
    chunk.changeLog = aggregateChangeLog(chunk);
    chunk.flags = aggregateFlags(chunk);
    moveToNextPending();
  }

  function approveCurrentChunk() {
    const chunk = currentChunk();
    if (!chunk) return;
    chunk.status = "approved";
    chunk.changeLog = aggregateChangeLog(chunk);
    chunk.flags = aggregateFlags(chunk);
    moveToNextPending();
  }

  function moveToNextPending() {
    const nextIndex = appState.chunks.findIndex((chunk) => chunk.status === "pending");
    if (nextIndex === -1) {
      appState.currentChunkIndex = appState.chunks.length - 1;
    } else {
      appState.currentChunkIndex = nextIndex;
    }
    scheduleSave();
    render();
  }

  function rechunkFromSource() {
    if (!appState.sourceParagraphs.length) return;
    appState.chunks = buildChunksFromParagraphs(appState.sourceParagraphs, appState.ui.chunkSizeTarget);
    appState.stagedParagraphs = [];
    appState.currentChunkIndex = 0;
    scheduleSave();
    render();
  }

  function exportTxt() {
    const approved = appState.chunks.filter((c) => c.status === "approved");
    if (!approved.length) {
      alert(getContent("alerts.noApprovedChunks", "No approved chunks to export."));
      return;
    }
    const text = approved
      .map((chunk) => chunk.paragraphs.map((p) => p.editedText).join("\n\n"))
      .join("\n\n");
    downloadBlob(text, `${baseFileName()}-reviewed.txt`, "text/plain");
  }

  function baseFileName() {
    if (!appState.meta.fileName) return "manuscript";
    return appState.meta.fileName.replace(/\.[^/.]+$/, "");
  }

  function downloadBlob(content, name, type) {
    const blob = content instanceof Blob ? content : new Blob([content], { type });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function exportSession() {
    downloadBlob(JSON.stringify(appState, null, 2), `${baseFileName()}-session.json`, "application/json");
  }

  function loadSession(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const parsed = JSON.parse(event.target.result);
        appState = hydrateState(parsed);
        scheduleSave();
        render();
      } catch (err) {
        alert(getContent("alerts.loadSessionError", "Unable to load session JSON."));
      }
    };
    reader.readAsText(file);
  }

  function exportDocx() {
    if (!window.docx) {
      alert(getContent("alerts.docxMissing", "docx.js is missing. Ensure docx.umd.js is loaded."));
      return;
    }
    const approved = appState.chunks.filter((c) => c.status === "approved");
    if (!approved.length) {
      alert(getContent("alerts.noApprovedChunks", "No approved chunks to export."));
      return;
    }

    const children = [];
    approved.forEach((chunk) => {
      chunk.paragraphs.forEach((para) => {
        children.push(buildDocxParagraph(para));
      });
    });

    const doc = new docx.Document({
      sections: [
        {
          properties: {
            page: {
              margin: buildDocxMargins(),
            },
          },
          children,
        },
      ],
    });

    docx.Packer.toBlob(doc).then((blob) => {
      downloadBlob(blob, `${baseFileName()}-reviewed.docx`, blob.type);
    });
  }

  function buildDocxMargins() {
    if (appState.docSettings.marginsPreset === "clean") {
      const marginCm = 2.5;
      const twip = cmToTwip(marginCm);
      return { top: twip, bottom: twip, left: twip, right: twip };
    }
    return { top: 1440, bottom: 1440, left: 1440, right: 1440 };
  }

  function buildDocxParagraph(para) {
    const settings = appState.docSettings;
    const baseSize = settings.fontSize * 2;
    const spacing = {
      before: ptToTwip(settings.spacingBeforePt),
      after: ptToTwip(settings.spacingAfterPt),
      line: Math.round(settings.lineSpacing * 240),
      lineRule: docx.LineRuleType.AUTO,
    };

    const indent = settings.indentEnabled
      ? { firstLine: cmToTwip(settings.firstLineIndentCm) }
      : {};

    const text = para.styleTag === "Poem" ? para.editedText : para.editedText.replace(/\n+/g, " ");

    if (para.styleTag === "Title") {
      return new docx.Paragraph({
        children: buildTextRuns(text, para.styleTag, settings.fontFamily, baseSize + 6),
        alignment: docx.AlignmentType.CENTER,
        spacing: { ...spacing, after: ptToTwip(12) },
      });
    }

    if (para.styleTag === "Heading1") {
      return new docx.Paragraph({
        children: buildTextRuns(text, para.styleTag, settings.fontFamily, baseSize + 2, { bold: true }),
        spacing: { ...spacing, before: ptToTwip(8), after: ptToTwip(6) },
      });
    }

    if (para.styleTag === "Quote") {
      return new docx.Paragraph({
        children: buildTextRuns(text, para.styleTag, settings.fontFamily, baseSize, { italics: true }),
        indent: { left: cmToTwip(0.8) },
        spacing,
      });
    }

    if (para.styleTag === "Poem") {
      return new docx.Paragraph({
        children: buildTextRuns(text, para.styleTag, settings.fontFamily, baseSize),
        spacing: { ...spacing, before: ptToTwip(2), after: ptToTwip(4) },
      });
    }

    return new docx.Paragraph({
      children: buildTextRuns(text, para.styleTag, settings.fontFamily, baseSize),
      spacing,
      indent,
    });
  }

  function buildTextRuns(text, styleTag, fontFamily, size, extra = {}) {
    const parts = styleTag === "Poem" ? text.split("\n") : [text];
    const runs = [];
    parts.forEach((part, index) => {
      runs.push(
        new docx.TextRun({
          text: part,
          font: fontFamily,
          size,
          ...extra,
          break: styleTag === "Poem" && index < parts.length - 1 ? 1 : 0,
        })
      );
    });
    return runs;
  }

  function cmToTwip(cm) {
    return Math.round(cm * 567);
  }

  function ptToTwip(pt) {
    return Math.round(pt * 20);
  }
  function bindEvents() {
    elements.fileInput.addEventListener("change", (event) => {
      const file = event.target.files[0];
      if (file) {
        pendingStoredState = null;
        elements.resumeBanner.classList.add("hidden");
        handleFileImport(file);
      }
      elements.fileInput.value = "";
    });

    elements.loadSessionBtn.addEventListener("click", () => {
      elements.sessionInput.click();
    });

    elements.sessionInput.addEventListener("change", (event) => {
      const file = event.target.files[0];
      if (file) {
        pendingStoredState = null;
        elements.resumeBanner.classList.add("hidden");
        loadSession(file);
      }
      elements.sessionInput.value = "";
    });

    elements.exportSessionBtn.addEventListener("click", exportSession);
    elements.exportTxtBtn.addEventListener("click", exportTxt);
    elements.exportDocxBtn.addEventListener("click", exportDocx);

    elements.docSettingsBtn.addEventListener("click", () => {
      elements.docSettingsModal.classList.remove("hidden");
    });

    elements.docSettingsClose.addEventListener("click", () => {
      elements.docSettingsModal.classList.add("hidden");
    });

    elements.chunkSize.addEventListener("input", (event) => {
      elements.chunkSizeValue.textContent = event.target.value;
      appState.ui.chunkSizeTarget = Number(event.target.value);
      scheduleSave();
    });

    elements.rechunkBtn.addEventListener("click", () => {
      if (!appState.sourceParagraphs.length) return;
      const confirmText = getContent("alerts.rechunkConfirm", "Rechunking will reset approvals and edits. Continue?");
      if (confirm(confirmText)) {
        rechunkFromSource();
      }
    });

    elements.layoutPreset.addEventListener("change", (event) => {
      appState.ui.layoutPreset = event.target.value;
      setLayoutPreset(appState.ui.layoutPreset);
      scheduleSave();
    });

    elements.engineSelect.addEventListener("change", (event) => {
      appState.ui.engineType = event.target.value;
      scheduleSave();
    });

    elements.modeSelect.addEventListener("change", (event) => {
      appState.ui.mode = event.target.value;
      scheduleSave();
    });

    elements.applyScope.addEventListener("change", (event) => {
      appState.ui.applyScope = event.target.value;
      scheduleSave();
    });

    elements.compareToggle.addEventListener("change", (event) => {
      appState.ui.compare = event.target.checked;
      scheduleSave();
      renderCurrentChunk();
    });

    if (elements.reviewPreviewToggle) {
      elements.reviewPreviewToggle.addEventListener("change", (event) => {
        appState.ui.reviewPreview = event.target.checked;
        scheduleSave();
        renderStaging();
      });
    }

    elements.analyzeBtn.addEventListener("click", runAnalyze);
    elements.suggestBtn.addEventListener("click", runSuggest);

    elements.resetChunkBtn.addEventListener("click", resetCurrentChunk);
    elements.skipChunkBtn.addEventListener("click", skipCurrentChunk);
    elements.approveChunkBtn.addEventListener("click", approveCurrentChunk);

    elements.resumeBtn.addEventListener("click", () => {
      if (!pendingStoredState) return;
      appState = hydrateState(pendingStoredState);
      pendingStoredState = null;
      elements.resumeBanner.classList.add("hidden");
      render();
    });

    elements.dismissBtn.addEventListener("click", () => {
      pendingStoredState = null;
      localStorage.removeItem(STORAGE_KEY);
      elements.resumeBanner.classList.add("hidden");
    });

    [
      elements.docFontFamily,
      elements.docFontSize,
      elements.docLineSpacing,
      elements.docMargins,
      elements.docIndentEnabled,
      elements.docIndentSize,
      elements.docSpaceBefore,
      elements.docSpaceAfter,
    ].forEach((input) => {
      input.addEventListener("change", () => {
        appState.docSettings.fontFamily = elements.docFontFamily.value;
        appState.docSettings.fontSize = Number(elements.docFontSize.value);
        appState.docSettings.lineSpacing = Number(elements.docLineSpacing.value);
        appState.docSettings.marginsPreset = elements.docMargins.value;
        appState.docSettings.indentEnabled = elements.docIndentEnabled.checked;
        appState.docSettings.firstLineIndentCm = Number(elements.docIndentSize.value);
        appState.docSettings.spacingBeforePt = Number(elements.docSpaceBefore.value);
        appState.docSettings.spacingAfterPt = Number(elements.docSpaceAfter.value);
        scheduleSave();
        renderStaging();
      });
    });

    document.addEventListener("keydown", (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key === "Enter") {
        event.preventDefault();
        approveCurrentChunk();
      }
      if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.key.toLowerCase() === "r") {
        event.preventDefault();
        resetCurrentChunk();
      }
    });
  }

  function init() {
    applyContent();
    loadFromStorage();
    render();
    bindEvents();
  }

  init();
})();
