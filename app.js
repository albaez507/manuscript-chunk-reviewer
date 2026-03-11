/*
Manuscript Chunk Reviewer
Run:
- Open index.html directly in a browser, or
- From this folder: python -m http.server 5173
*/

(() => {
  const STORAGE_KEY = "mcr_state_v1";
  const EMPTY_STATE_ILLUSTRATION = "assets/illustrations/empty-state.svg";
  const CONTENT = window.MCR_CONTENT || {};
  const STYLE_EXPORT_MAP = {
    Normal: "Normal",
    Title: "Title",
    Heading1: "Heading1",
    Quote: "Quote",
    Poem: "Poem",
    "Book Title": "Title",
    "Book Subtitle": "Heading1",
    "Author Name": "Heading1",
    "Part Title": "Heading1",
    "Part Subtitle": "Heading1",
    "Page Title": "Heading1",
    "First Paragraph": "Normal",
    Dedication: "Normal",
    "Opening Quote": "Quote",
    "Opening Quote Credit": "Normal",
    "Copyright Text": "Normal",
    "Chapter Title": "Heading1",
    "Chapter Subtitle": "Heading1",
    "Chapter First Paragraph": "Normal",
    Separator: "Normal",
    "Block Quote": "Quote",
  };
  const STYLE_LABEL_MAP = {
    Normal: "Normal",
    Title: "Book Title",
    Heading1: "Chapter Title",
    Quote: "Block Quote",
    Poem: "Poem",
  };

  const $ = (id) => document.getElementById(id);

  const elements = {
    setupScreen: $("setup-screen"),
    editorShell: $("editor-shell"),
    projectTitle: $("project-title"),
    projectAuthor: $("project-author"),
    projectLanguage: $("project-language"),
    projectType: $("project-type"),
    projectFileInput: $("project-file-input"),
    projectTrimSize: $("project-trim-size"),
    projectBleed: $("project-bleed"),
    projectMarginTop: $("project-margin-top"),
    projectMarginBottom: $("project-margin-bottom"),
    projectMarginInside: $("project-margin-inside"),
    projectMarginOutside: $("project-margin-outside"),
    projectGutter: $("project-gutter"),
    projectMirror: $("project-mirror"),
    projectPageSide: $("project-page-side"),
    projectHeaderText: $("project-header-text"),
    projectFooterText: $("project-footer-text"),
    projectPageNumberPosition: $("project-page-number-position"),
    projectFontFamily: $("project-font-family"),
    projectFontSize: $("project-font-size"),
    projectLineSpacing: $("project-line-spacing"),
    projectIndentSize: $("project-indent-size"),
    projectSpaceBefore: $("project-space-before"),
    projectSpaceAfter: $("project-space-after"),
    projectStartBtn: $("project-start-btn"),
    projectValidation: $("project-validation"),
    paginateBtn: $("paginate-btn"),
    projectSettingsBtn: $("project-settings-btn"),
    outlineFront: $("outline-front"),
    outlineBody: $("outline-body"),
    outlineBack: $("outline-back"),
    addFrontBtn: $("add-front-btn"),
    addBodyBtn: $("add-body-btn"),
    addBackBtn: $("add-back-btn"),
    resumeBanner: $("resume-banner"),
    resumeMeta: $("resume-meta"),
    resumeBtn: $("resume-btn"),
    dismissBtn: $("dismiss-btn"),
    fileMeta: $("file-meta"),
    sessionInput: $("session-input"),
    loadSessionBtn: $("load-session-btn"),
    exportSessionBtn: $("export-session-btn"),
    exportTxtBtn: $("export-txt-btn"),
    exportDocxBtn: $("export-docx-btn"),
    docSettingsBtn: $("doc-settings-btn"),
    docSettingsModal: $("doc-settings-modal"),
    docSettingsClose: $("doc-settings-close"),
    pageInfo: $("page-info"),
    pageStats: $("page-stats"),
    prevPageBtn: $("prev-page-btn"),
    nextPageBtn: $("next-page-btn"),
    pageSection: $("page-section"),
    pageTitle: $("page-title"),
    editorBody: $("editor-body"),
    splitter: $("splitter"),
    propertiesTabElements: $("properties-tab-elements"),
    propertiesTabFormatting: $("properties-tab-formatting"),
    propertiesElementsSection: $("properties-elements"),
    propertiesFormattingSection: $("properties-formatting"),
    propertiesCurrent: $("properties-current"),
    propertiesClearBtn: $("properties-clear"),
    formatFontFamily: $("format-font-family"),
    formatFontSize: $("format-font-size"),
    formatLineSpacing: $("format-line-spacing"),
    formatAlign: $("format-align"),
    formatIndent: $("format-indent"),
    formatSpacingBefore: $("format-spacing-before"),
    formatSpacingAfter: $("format-spacing-after"),
    formatDropCap: $("format-drop-cap"),
    docFontFamily: $("doc-font-family"),
    docFontSize: $("doc-font-size"),
    docLineSpacing: $("doc-line-spacing"),
    docMargins: $("doc-margins"),
    docTrimSize: $("doc-trim-size"),
    docPageSide: $("doc-page-side"),
    docHeaderText: $("doc-header-text"),
    docFooterText: $("doc-footer-text"),
    docPageNumberPosition: $("doc-page-number-position"),
    docMarginTop: $("doc-margin-top"),
    docMarginBottom: $("doc-margin-bottom"),
    docMarginInside: $("doc-margin-inside"),
    docMarginOutside: $("doc-margin-outside"),
    docGutter: $("doc-gutter"),
    docMirrorMargins: $("doc-mirror-margins"),
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
      project: {
        title: "",
        author: "",
        language: "en",
        type: "print",
        bleed: false,
        pageCountOverride: null,
      },
      ui: {
        view: "setup",
      },
      docSettings: {
        trimSize: "6x9",
        pageSide: "auto",
        headerText: "",
        footerText: "",
        pageNumberPosition: "footer-center",
        fontFamily: "Times New Roman",
        fontSize: 12,
        lineSpacing: 1.2,
        textAlign: "justify",
        marginsPreset: "clean",
        marginTopIn: 0.75,
        marginBottomIn: 0.75,
        marginInsideIn: 0.75,
        marginOutsideIn: 0.5,
        gutterIn: 0.13,
        mirrorMargins: true,
        indentEnabled: true,
        firstLineIndentCm: 0.5,
        spacingBeforePt: 0,
        spacingAfterPt: 6,
      },
      sourceParagraphs: [],
      sourceText: "",
      pages: [],
      currentPageIndex: 0,
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
    const legacyPages = Array.isArray(state.chunks) ? state.chunks : [];
    const pages = Array.isArray(state.pages) ? state.pages : legacyPages;
    const merged = {
      ...defaults,
      ...state,
      meta: { ...defaults.meta, ...(state.meta || {}) },
      project: { ...defaults.project, ...(state.project || {}) },
      ui: { ...defaults.ui, ...(state.ui || {}) },
      docSettings: { ...defaults.docSettings, ...(state.docSettings || {}) },
      sourceParagraphs: Array.isArray(state.sourceParagraphs) ? state.sourceParagraphs : [],
      sourceText: typeof state.sourceText === "string" ? state.sourceText : "",
      pages,
      currentPageIndex: Number.isInteger(state.currentPageIndex)
        ? state.currentPageIndex
        : Number.isInteger(state.currentChunkIndex)
          ? state.currentChunkIndex
          : 0,
    };

    merged.pages.forEach((page, pageIndex) => {
      if (!Array.isArray(page.paragraphs)) {
        page.paragraphs = [];
      }
      if (!page.section) {
        page.section = "body";
      }
      if (!page.title) {
        page.title = `Page ${pageIndex + 1}`;
      }
      page.paragraphs.forEach((para) => {
        if (!para.id) {
          para.id = uid("p");
        }
        if (typeof para.originalText !== "string") {
          para.originalText = "";
        }
        if (typeof para.editedText !== "string") {
          para.editedText = para.originalText;
        }
        if (!para.styleTag) {
          para.styleTag = "Normal";
        }
        if (para.format && typeof para.format !== "object") {
          delete para.format;
        }
        delete para.suggestedText;
        delete para.reviewDecision;
        delete para.changeLog;
        delete para.flags;
        ensureHistory(para);
      });
    });

    return merged;
  }

  function renderFileMeta() {
    if (!elements.fileMeta) return;
    if (appState.meta.fileName) {
      const date = appState.meta.importedAt
        ? new Date(appState.meta.importedAt).toLocaleString()
        : "";
      elements.fileMeta.textContent = `${appState.meta.fileName} | ${appState.meta.inputFormat.toUpperCase()} | ${date}`;
      return;
    }
    const title = appState.project?.title || "";
    const author = appState.project?.author || "";
    if (title || author) {
      elements.fileMeta.textContent = [title, author].filter(Boolean).join(" | ");
      return;
    }
    elements.fileMeta.textContent = getContent("brand.subtitleNoFile", "No file loaded");
  }

  function parseParagraphs(text) {
    if (!text) return [];
    return text
      .replace(/\r\n/g, "\n")
      .split(/\n\s*\n+/)
      .map((para) => para.trim())
      .filter(Boolean);
  }

  function estimateLinesForText(text, maxCharsPerLine) {
    const clean = (text || "").trim();
    if (!clean) return 1;
    const words = clean.split(/\s+/);
    let lines = 1;
    let current = 0;
    words.forEach((word) => {
      const next = current ? current + 1 + word.length : word.length;
      if (next > maxCharsPerLine) {
        lines += 1;
        current = word.length;
      } else {
        current = next;
      }
    });
    return Math.max(1, lines);
  }

  function splitParagraphByLines(text, maxCharsPerLine, maxLines) {
    const words = (text || "").trim().split(/\s+/).filter(Boolean);
    if (!words.length) return [""];
    const parts = [];
    let current = [];
    words.forEach((word) => {
      const next = current.length ? [...current, word].join(" ") : word;
      if (estimateLinesForText(next, maxCharsPerLine) > maxLines && current.length) {
        parts.push(current.join(" "));
        current = [word];
      } else {
        current.push(word);
      }
    });
    if (current.length) {
      parts.push(current.join(" "));
    }
    return parts;
  }

  function calculatePageMetrics(settings) {
    const trim = getTrimSize(settings.trimSize);
    const dpi = 96;
    const fontSizePx = Math.round(settings.fontSize * 1.333);
    const lineHeightPx = Math.round(fontSizePx * settings.lineSpacing);
    const pageWidthPx = inchesToPx(trim.width, dpi);
    const pageHeightPx = inchesToPx(trim.height, dpi);
    const gutterPx = inchesToPx(settings.gutterIn || 0, dpi);
    const leftPx = inchesToPx(settings.marginInsideIn || 0.75, dpi) + gutterPx;
    const rightPx = inchesToPx(settings.marginOutsideIn || 0.5, dpi);
    const topPx = inchesToPx(settings.marginTopIn || 0.75, dpi);
    const bottomPx = inchesToPx(settings.marginBottomIn || 0.75, dpi);

    const headerFooterLines =
      settings.headerText || settings.footerText || settings.pageNumberPosition !== "none" ? 2 : 0;
    const contentWidth = Math.max(120, pageWidthPx - leftPx - rightPx);
    const contentHeight = Math.max(120, pageHeightPx - topPx - bottomPx - headerFooterLines * lineHeightPx);
    const maxCharsPerLine = Math.max(20, Math.floor(contentWidth / (fontSizePx * 0.52)));
    const maxLines = Math.max(5, Math.floor(contentHeight / lineHeightPx));
    const spacingBeforePx = Math.round(settings.spacingBeforePt * 1.333);
    const spacingAfterPx = Math.round(settings.spacingAfterPt * 1.333);
    const spacingLines = Math.ceil((spacingBeforePx + spacingAfterPx) / lineHeightPx);

    return { maxCharsPerLine, maxLines, spacingLines };
  }

  function buildPagesFromParagraphArrays(pages) {
    return pages.map((pageParas, pageIndex) => ({
      id: uid("page"),
      section: "body",
      title: `Page ${pageIndex + 1}`,
      paragraphs: pageParas.map((text) => createParagraphFromText(text)),
    }));
  }

  function paginateFromParagraphs(paragraphs, settings) {
    const { maxCharsPerLine, maxLines, spacingLines } = calculatePageMetrics(settings);
    const pages = [];
    let current = [];
    let currentLines = 0;

    paragraphs.forEach((para) => {
      const estimatedLines = estimateLinesForText(para, maxCharsPerLine) + spacingLines;
      if (estimatedLines > maxLines) {
        const splits = splitParagraphByLines(para, maxCharsPerLine, Math.max(1, maxLines - spacingLines));
        splits.forEach((part) => {
          const partLines = estimateLinesForText(part, maxCharsPerLine) + spacingLines;
          if (currentLines + partLines > maxLines && current.length) {
            pages.push(current);
            current = [];
            currentLines = 0;
          }
          current.push(part);
          currentLines += partLines;
        });
        return;
      }
      if (currentLines + estimatedLines > maxLines && current.length) {
        pages.push(current);
        current = [];
        currentLines = 0;
      }
      current.push(para);
      currentLines += estimatedLines;
    });

    if (current.length) {
      pages.push(current);
    }
    return buildPagesFromParagraphArrays(pages);
  }

  function collectEditedText() {
    if (!appState.pages || !appState.pages.length) {
      return appState.sourceText || "";
    }
    return appState.pages
      .map((page) => page.paragraphs.map((p) => p.editedText).join("\n\n"))
      .join("\n\n");
  }

  function repaginateFromText(text) {
    const paragraphs = parseParagraphs(text);
    appState.sourceText = text;
    appState.sourceParagraphs = paragraphs;
    appState.pages = paginateFromParagraphs(paragraphs, appState.docSettings);
    appState.currentPageIndex = 0;
    activeParagraphId = null;
  }

  function ensureAtLeastOnePage() {
    if (appState.pages.length) return;
    appState.pages = buildPagesFromParagraphArrays([[""]]);
    appState.currentPageIndex = 0;
  }

  function validatePrintSettings(settings, pageCount, bleed) {
    const issues = [];
    const minOuter = bleed ? 0.375 : 0.25;

    if (settings.marginTopIn < minOuter) {
      issues.push(`Top margin must be >= ${minOuter}"`);
    }
    if (settings.marginBottomIn < minOuter) {
      issues.push(`Bottom margin must be >= ${minOuter}"`);
    }
    if (settings.marginOutsideIn < minOuter) {
      issues.push(`Outside margin must be >= ${minOuter}"`);
    }

    const count = pageCount || 24;
    const countLabel = pageCount ? `${pageCount} pages` : "estimated pages";
    let minInside = 0.375;
    if (count >= 151 && count <= 300) minInside = 0.5;
    if (count >= 301 && count <= 500) minInside = 0.625;
    if (count >= 501 && count <= 700) minInside = 0.75;
    if (count >= 701) minInside = 0.875;

    if (settings.marginInsideIn < minInside) {
      issues.push(`Inside margin must be >= ${minInside}" for ${countLabel}`);
    }
    if (settings.gutterIn < minInside) {
      issues.push(`Gutter must be >= ${minInside}" for ${countLabel}`);
    }
    if (Math.abs(settings.marginInsideIn - settings.gutterIn) > 0.01) {
      issues.push("Inside margin and gutter should match for KDP Print.");
    }
    if (pageCount && pageCount < 24) {
      issues.push("KDP Print requires a minimum of 24 pages.");
    }
    return issues;
  }

  function createParagraphFromText(text) {
    const safeText = text ?? "";
    return {
      id: uid("p"),
      originalText: safeText,
      editedText: safeText,
      styleTag: "Normal",
      history: {
        stack: [safeText],
        index: 0,
      },
    };
  }

  function createPage(section, paragraphTexts = [""]) {
    return {
      id: uid("page"),
      section: section || "body",
      title: "",
      paragraphs: paragraphTexts.map((text) => createParagraphFromText(text)),
    };
  }

  function cloneParagraph(para) {
    const text = para?.editedText ?? para?.originalText ?? "";
    const copy = createParagraphFromText(text);
    copy.styleTag = para?.styleTag || "Normal";
    if (para?.format && typeof para.format === "object") {
      copy.format = { ...para.format };
    }
    return copy;
  }

  function clonePage(page) {
    return {
      id: uid("page"),
      section: page?.section || "body",
      title: page?.title || "",
      paragraphs: (page?.paragraphs || []).map((para) => cloneParagraph(para)),
    };
  }

  function findInsertIndexForSection(section) {
    const pages = appState.pages || [];
    if (!pages.length) return 0;
    let lastIndex = -1;
    pages.forEach((page, index) => {
      if (page.section === section) {
        lastIndex = index;
      }
    });
    if (lastIndex !== -1) return lastIndex + 1;
    if (section === "front") return 0;
    if (section === "back") return pages.length;
    const firstBack = pages.findIndex((page) => page.section === "back");
    return firstBack === -1 ? pages.length : firstBack;
  }

  function insertPageAt(index, page) {
    const bounded = Math.max(0, Math.min(index, appState.pages.length));
    appState.pages.splice(bounded, 0, page);
    appState.currentPageIndex = bounded;
    activeParagraphId = null;
    scheduleSave();
    render();
  }

  function addBlankPage(section) {
    const page = createPage(section, [""]);
    const insertIndex = findInsertIndexForSection(section);
    insertPageAt(insertIndex, page);
  }

  function duplicatePageAt(index) {
    const source = appState.pages[index];
    if (!source) return;
    const copy = clonePage(source);
    insertPageAt(index + 1, copy);
  }

  function deletePageAt(index) {
    const pages = appState.pages || [];
    if (!pages.length) return;
    if (pages.length === 1) {
      const only = pages[0];
      only.title = "";
      only.section = only.section || "body";
      only.paragraphs = [createParagraphFromText("")];
      activeParagraphId = null;
      scheduleSave();
      render();
      return;
    }

    pages.splice(index, 1);
    if (appState.currentPageIndex >= pages.length) {
      appState.currentPageIndex = pages.length - 1;
    } else if (index <= appState.currentPageIndex) {
      appState.currentPageIndex = Math.max(0, appState.currentPageIndex - 1);
    }
    activeParagraphId = null;
    scheduleSave();
    render();
  }

  function movePageToIndex(sourceIndex, targetIndex, targetSection) {
    const pages = appState.pages || [];
    if (sourceIndex < 0 || sourceIndex >= pages.length) return;
    if (targetIndex < 0) targetIndex = 0;
    if (targetIndex > pages.length) targetIndex = pages.length;
    const [moved] = pages.splice(sourceIndex, 1);
    moved.section = targetSection || moved.section || "body";
    let insertIndex = targetIndex;
    if (sourceIndex < targetIndex) {
      insertIndex = targetIndex - 1;
    }
    insertIndex = Math.max(0, Math.min(insertIndex, pages.length));
    pages.splice(insertIndex, 0, moved);
    appState.currentPageIndex = insertIndex;
    activeParagraphId = null;
    scheduleSave();
    render();
  }

  function buildOutlineAction(label, handler) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "outline-action-btn";
    button.textContent = label;
    button.addEventListener("click", (event) => {
      event.stopPropagation();
      handler();
    });
    button.addEventListener("dragstart", (event) => event.preventDefault());
    return button;
  }

  function renamePageAt(index) {
    const page = appState.pages[index];
    if (!page) return;
    const fallback = `Page ${index + 1}`;
    const nextTitle = prompt(getContent("prompts.renamePage", "Page title:"), page.title || fallback);
    if (nextTitle === null) return;
    page.title = nextTitle.trim();
    scheduleSave();
    renderOutline();
    if (index === appState.currentPageIndex) {
      renderCurrentPage();
    }
  }

  function wireOutlineList(listEl, section) {
    if (!listEl || listEl.dataset.dndBound === "true") return;
    listEl.dataset.dndBound = "true";
    listEl.dataset.section = section;
    listEl.addEventListener("dragover", (event) => {
      event.preventDefault();
      listEl.classList.add("drag-over");
      if (event.dataTransfer) {
        event.dataTransfer.dropEffect = "move";
      }
    });
    listEl.addEventListener("dragleave", () => {
      listEl.classList.remove("drag-over");
    });
    listEl.addEventListener("drop", (event) => {
      event.preventDefault();
      listEl.classList.remove("drag-over");
      const sourceIndex = Number(event.dataTransfer?.getData("text/plain"));
      if (Number.isNaN(sourceIndex)) return;
      const targetIndex = findInsertIndexForSection(section);
      movePageToIndex(sourceIndex, targetIndex, section);
    });
  }

  function currentPage() {
    return appState.pages[appState.currentPageIndex] || null;
  }

  function calcTextStats(text) {
    const trimmed = text.trim();
    const words = trimmed ? trimmed.split(/\s+/).length : 0;
    const chars = trimmed.length;
    return { words, chars };
  }

  function calcPageStats(page) {
    if (!page) return { words: 0, chars: 0 };
    const text = page.paragraphs.map((p) => p.editedText).join("\n\n");
    return calcTextStats(text);
  }

  function renderSetup() {
    if (!elements.setupScreen) return;
    elements.setupScreen.classList.remove("hidden");
    if (elements.editorShell) {
      elements.editorShell.classList.add("hidden");
    }

    if (elements.projectTitle) elements.projectTitle.value = appState.project.title || "";
    if (elements.projectAuthor) elements.projectAuthor.value = appState.project.author || "";
    if (elements.projectLanguage) elements.projectLanguage.value = appState.project.language || "en";
    if (elements.projectType) elements.projectType.value = appState.project.type || "print";
    if (elements.projectTrimSize) elements.projectTrimSize.value = appState.docSettings.trimSize;
    if (elements.projectBleed) elements.projectBleed.value = appState.project.bleed ? "on" : "off";
    if (elements.projectMarginTop) elements.projectMarginTop.value = appState.docSettings.marginTopIn;
    if (elements.projectMarginBottom) elements.projectMarginBottom.value = appState.docSettings.marginBottomIn;
    if (elements.projectMarginInside) elements.projectMarginInside.value = appState.docSettings.marginInsideIn;
    if (elements.projectMarginOutside) elements.projectMarginOutside.value = appState.docSettings.marginOutsideIn;
    if (elements.projectGutter) elements.projectGutter.value = appState.docSettings.gutterIn;
    if (elements.projectMirror) elements.projectMirror.value = appState.docSettings.mirrorMargins ? "on" : "off";
    if (elements.projectPageSide) elements.projectPageSide.value = appState.docSettings.pageSide;
    if (elements.projectHeaderText) elements.projectHeaderText.value = appState.docSettings.headerText;
    if (elements.projectFooterText) elements.projectFooterText.value = appState.docSettings.footerText;
    if (elements.projectPageNumberPosition) {
      elements.projectPageNumberPosition.value = appState.docSettings.pageNumberPosition;
    }
    if (elements.projectFontFamily) elements.projectFontFamily.value = appState.docSettings.fontFamily;
    if (elements.projectFontSize) elements.projectFontSize.value = appState.docSettings.fontSize;
    if (elements.projectLineSpacing) elements.projectLineSpacing.value = appState.docSettings.lineSpacing;
    if (elements.projectIndentSize) elements.projectIndentSize.value = appState.docSettings.firstLineIndentCm;
    if (elements.projectSpaceBefore) elements.projectSpaceBefore.value = appState.docSettings.spacingBeforePt;
    if (elements.projectSpaceAfter) elements.projectSpaceAfter.value = appState.docSettings.spacingAfterPt;

    const printSection = document.getElementById("project-print-settings");
    if (printSection) {
      const isPrint = (appState.project.type || "print") === "print";
      printSection.classList.toggle("disabled", !isPrint);
    }

    const pageCount = appState.pages.length;
    if (elements.projectValidation) {
      if ((appState.project.type || "print") === "print") {
        const issues = validatePrintSettings(appState.docSettings, pageCount, appState.project.bleed);
        elements.projectValidation.textContent = issues.length
          ? `Validation: ${issues.join(" | ")}`
          : "Validation: OK";
      } else {
        elements.projectValidation.textContent = "Validation: eBook selected (print margins ignored).";
      }
    }
  }

  function applyProjectSettingsFromSetup() {
    appState.project.title = elements.projectTitle?.value?.trim() || "";
    appState.project.author = elements.projectAuthor?.value?.trim() || "";
    appState.project.language = elements.projectLanguage?.value || "en";
    appState.project.type = elements.projectType?.value || "print";
    appState.project.bleed = (elements.projectBleed?.value || "off") === "on";

    appState.docSettings.trimSize = elements.projectTrimSize?.value || appState.docSettings.trimSize;
    appState.docSettings.marginTopIn = Number(elements.projectMarginTop?.value || appState.docSettings.marginTopIn);
    appState.docSettings.marginBottomIn = Number(
      elements.projectMarginBottom?.value || appState.docSettings.marginBottomIn
    );
    appState.docSettings.marginInsideIn = Number(
      elements.projectMarginInside?.value || appState.docSettings.marginInsideIn
    );
    appState.docSettings.marginOutsideIn = Number(
      elements.projectMarginOutside?.value || appState.docSettings.marginOutsideIn
    );
    appState.docSettings.gutterIn = Number(elements.projectGutter?.value || appState.docSettings.gutterIn);
    appState.docSettings.mirrorMargins = (elements.projectMirror?.value || "on") === "on";
    appState.docSettings.pageSide = elements.projectPageSide?.value || appState.docSettings.pageSide;
    appState.docSettings.headerText = elements.projectHeaderText?.value || "";
    appState.docSettings.footerText = elements.projectFooterText?.value || "";
    appState.docSettings.pageNumberPosition =
      elements.projectPageNumberPosition?.value || appState.docSettings.pageNumberPosition;
    appState.docSettings.fontFamily = elements.projectFontFamily?.value || appState.docSettings.fontFamily;
    appState.docSettings.fontSize = Number(elements.projectFontSize?.value || appState.docSettings.fontSize);
    appState.docSettings.lineSpacing = Number(
      elements.projectLineSpacing?.value || appState.docSettings.lineSpacing
    );
    appState.docSettings.firstLineIndentCm = Number(
      elements.projectIndentSize?.value || appState.docSettings.firstLineIndentCm
    );
    appState.docSettings.spacingBeforePt = Number(
      elements.projectSpaceBefore?.value || appState.docSettings.spacingBeforePt
    );
    appState.docSettings.spacingAfterPt = Number(
      elements.projectSpaceAfter?.value || appState.docSettings.spacingAfterPt
    );

  }

  function startProject() {
    applyProjectSettingsFromSetup();
    const text = appState.sourceText || "";
    if (text) {
      repaginateFromText(text);
    } else {
      appState.sourceText = "";
      appState.sourceParagraphs = [];
      appState.pages = [];
      ensureAtLeastOnePage();
    }
    appState.ui.view = "editor";
    scheduleSave();
    render();
  }

  function openProjectSettings() {
    appState.sourceText = collectEditedText();
    appState.ui.view = "setup";
    scheduleSave();
    render();
  }

  function renderToolbar() {
    if (elements.docTrimSize) {
      elements.docTrimSize.value = appState.docSettings.trimSize;
    }
    if (elements.docPageSide) {
      elements.docPageSide.value = appState.docSettings.pageSide;
    }
    if (elements.docHeaderText) {
      elements.docHeaderText.value = appState.docSettings.headerText;
    }
    if (elements.docFooterText) {
      elements.docFooterText.value = appState.docSettings.footerText;
    }
    if (elements.docPageNumberPosition) {
      elements.docPageNumberPosition.value = appState.docSettings.pageNumberPosition;
    }
    if (elements.docFontFamily) elements.docFontFamily.value = appState.docSettings.fontFamily;
    if (elements.docFontSize) elements.docFontSize.value = appState.docSettings.fontSize;
    if (elements.docLineSpacing) elements.docLineSpacing.value = appState.docSettings.lineSpacing;
    if (elements.docMargins) elements.docMargins.value = appState.docSettings.marginsPreset;
    if (elements.docMarginTop) {
      elements.docMarginTop.value = appState.docSettings.marginTopIn;
    }
    if (elements.docMarginBottom) {
      elements.docMarginBottom.value = appState.docSettings.marginBottomIn;
    }
    if (elements.docMarginInside) {
      elements.docMarginInside.value = appState.docSettings.marginInsideIn;
    }
    if (elements.docMarginOutside) {
      elements.docMarginOutside.value = appState.docSettings.marginOutsideIn;
    }
    if (elements.docGutter) {
      elements.docGutter.value = appState.docSettings.gutterIn;
    }
    if (elements.docMirrorMargins) {
      elements.docMirrorMargins.checked = appState.docSettings.mirrorMargins;
    }
    if (elements.docIndentEnabled) elements.docIndentEnabled.checked = appState.docSettings.indentEnabled;
    if (elements.docIndentSize) elements.docIndentSize.value = appState.docSettings.firstLineIndentCm;
    if (elements.docSpaceBefore) elements.docSpaceBefore.value = appState.docSettings.spacingBeforePt;
    if (elements.docSpaceAfter) elements.docSpaceAfter.value = appState.docSettings.spacingAfterPt;
  }

  function renderCurrentPage() {
    const page = currentPage();
    if (!page) {
      elements.pageInfo.textContent = getContent("messages.noPages", "No pages loaded");
      elements.pageStats.textContent = "";
      const emptyText = getContent("messages.emptyState", "Import a file to begin page layout.");
      elements.editorBody.innerHTML = `
        <div class="empty-state">
          <img src="${EMPTY_STATE_ILLUSTRATION}" alt="" aria-hidden="true" />
          <div>${emptyText}</div>
        </div>
      `;
      return;
    }
    if (!activeParagraphId || !page.paragraphs.some((para) => para.id === activeParagraphId)) {
      activeParagraphId = page.paragraphs[0]?.id || null;
    }

    const pageIndex = appState.currentPageIndex + 1;
    const stats = calcPageStats(page);
    const pageLabel = getContent("labels.pageLabel", "Page");
    const totalPages = appState.pages.length;
    const titleText = page.title ? ` | ${page.title}` : "";
    elements.pageInfo.textContent = `${pageLabel} ${pageIndex} of ${totalPages}${titleText}`;
    elements.pageStats.textContent = `${stats.words} words | ${stats.chars} chars`;
    if (elements.pageSection) {
      elements.pageSection.value = page.section || "body";
    }
    if (elements.pageTitle) {
      elements.pageTitle.value = page.title || `Page ${pageIndex}`;
    }
    if (elements.prevPageBtn) {
      elements.prevPageBtn.disabled = appState.currentPageIndex <= 0;
    }
    if (elements.nextPageBtn) {
      elements.nextPageBtn.disabled = appState.currentPageIndex >= appState.pages.length - 1;
    }

    elements.editorBody.innerHTML = "";

    const pageNumber = appState.currentPageIndex + 1;
    const hostWidth = elements.editorBody?.clientWidth || 720;
    const { frame, body: pageBody } = buildPageShell(appState.docSettings, pageNumber, hostWidth, "editor-page");
    frame.classList.add("editor-frame");
    page.paragraphs.forEach((para) => {
      ensureHistory(para);
      const format = getParagraphFormat(para);
      const indentPx = format.indentEnabled ? cmToPx(format.indentCm || 0) : 0;
      const spacingBeforePx = ptToPx(format.spacingBeforePt || 0);
      const spacingAfterPx = ptToPx(format.spacingAfterPt || 0);
      const block = document.createElement("div");
      block.className = "paragraph-block";
      block.style.marginTop = `${spacingBeforePx}px`;
      block.style.marginBottom = `${spacingAfterPx}px`;
      if (para.id === activeParagraphId) {
        block.classList.add("active");
      }

      const body = document.createElement("div");
      body.className = "paragraph-body";

      const edited = document.createElement("textarea");
      edited.value = para.editedText;
      applyEditorParagraphStyles(edited, {
        indentPx,
        textAlign: format.textAlign,
        lineHeight: String(format.lineSpacing || 1.2),
        fontFamily: format.fontFamily,
        fontSizePx: Math.round(format.fontSize * 1.333),
      });
      body.appendChild(edited);

      edited.addEventListener("focus", () => {
        activeParagraphId = para.id;
        updatePropertiesPanel();
      });
      edited.addEventListener("input", (event) => {
        para.editedText = event.target.value;
        autoSizeTextarea(event.target);
        scheduleHistory(para, para.editedText);
        scheduleSave();
        renderCurrentPageStatsOnly();
      });
      edited.addEventListener("blur", () => {
        recordHistory(para, para.editedText);
      });
      edited.addEventListener("keydown", (event) => {
        const key = event.key.toLowerCase();
        if ((event.ctrlKey || event.metaKey) && !event.shiftKey && key === "z") {
          event.preventDefault();
          undoParagraph(para, edited);
        }
        if ((event.ctrlKey || event.metaKey) && (key === "y" || (event.shiftKey && key === "z"))) {
          event.preventDefault();
          redoParagraph(para, edited);
        }
      });
      autoSizeTextarea(edited);

      block.appendChild(body);
      pageBody.appendChild(block);
    });

    elements.editorBody.appendChild(frame);
  }

  function findParagraphLocation(paragraphId, pageId) {
    const pages = appState.pages || [];
    if (pageId) {
      const pageIndex = pages.findIndex((page) => page.id === pageId);
      if (pageIndex >= 0) {
        const page = pages[pageIndex];
        const paraIndex = page.paragraphs.findIndex((p) => p.id === paragraphId);
        if (paraIndex >= 0) {
          return { page, pageIndex, paragraph: page.paragraphs[paraIndex], paraIndex };
        }
      }
    }
    for (let i = 0; i < pages.length; i += 1) {
      const page = pages[i];
      const paraIndex = page.paragraphs.findIndex((p) => p.id === paragraphId);
      if (paraIndex >= 0) {
        return { page, pageIndex: i, paragraph: page.paragraphs[paraIndex], paraIndex };
      }
    }
    return null;
  }

  function jumpToParagraph(location) {
    if (!location) return;
    appState.currentPageIndex = location.pageIndex;
    activeParagraphId = location.paragraph.id;
    scheduleSave();
    render();
  }

  function importText(text, meta) {
    const paragraphs = parseParagraphs(text);
    appState.meta = { ...appState.meta, ...meta };
    appState.sourceText = text;
    appState.sourceParagraphs = paragraphs;
    scheduleSave();
    render();
  }

  async function extractDocxTextFallback(arrayBuffer) {
    if (!window.JSZip) return "";
    const zip = await window.JSZip.loadAsync(arrayBuffer);
    const documentXml = await zip.file("word/document.xml")?.async("string");
    if (!documentXml) return "";
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(documentXml, "application/xml");
    const textNodes = Array.from(xmlDoc.getElementsByTagName("w:t"));
    return textNodes.map((node) => node.textContent || "").join(" ");
  }

  function handleFileImport(file) {
    if (!file) return;
    if (file.size === 0) {
      alert(getContent("alerts.emptyFile", "The file is empty."));
      return;
    }
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
          .catch(async (err) => {
            console.warn("Docx parse error", err);
            try {
              const fallbackText = await extractDocxTextFallback(arrayBuffer);
              if (fallbackText) {
                importText(fallbackText, meta);
                return;
              }
            } catch (fallbackErr) {
              console.warn("Docx fallback error", fallbackErr);
            }
            const base = getContent("alerts.docxParseError", "Unable to parse .docx file.");
            const detail = err && err.message ? ` (${err.message})` : "";
            const fallbackMsg = getContent(
              "alerts.docxFallbackError",
              "Unable to read this .docx. Try saving it again as a Word .docx file."
            );
            alert(`${base}${detail}\\n${fallbackMsg}`);
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

  function renderOutline() {
    if (!elements.outlineFront || !elements.outlineBody || !elements.outlineBack) return;
    elements.outlineFront.innerHTML = "";
    elements.outlineBody.innerHTML = "";
    elements.outlineBack.innerHTML = "";

    wireOutlineList(elements.outlineFront, "front");
    wireOutlineList(elements.outlineBody, "body");
    wireOutlineList(elements.outlineBack, "back");

    appState.pages.forEach((page, index) => {
      const item = document.createElement("div");
      item.className = "outline-item";
      item.draggable = true;
      if (index === appState.currentPageIndex) {
        item.classList.add("active");
      }

      const left = document.createElement("div");
      left.className = "outline-left";

      const title = page.title || `Page ${index + 1}`;
      const label = document.createElement("div");
      label.className = "outline-label";
      label.textContent = `${index + 1}. ${title}`;

      left.appendChild(label);

      const actions = document.createElement("div");
      actions.className = "outline-item-actions";
      actions.appendChild(buildOutlineAction("Rename", () => renamePageAt(index)));
      actions.appendChild(buildOutlineAction("Duplicate", () => duplicatePageAt(index)));
      actions.appendChild(
        buildOutlineAction("Delete", () => {
          const confirmText = getContent("alerts.deletePageConfirm", "Delete this page?");
          if (confirm(confirmText)) {
            deletePageAt(index);
          }
        })
      );

      item.appendChild(left);
      item.appendChild(actions);
      item.addEventListener("click", () => {
        appState.currentPageIndex = index;
        activeParagraphId = null;
        scheduleSave();
        render();
      });

      item.addEventListener("dragstart", (event) => {
        item.classList.add("dragging");
        if (event.dataTransfer) {
          event.dataTransfer.effectAllowed = "move";
          event.dataTransfer.setData("text/plain", String(index));
        }
      });
      item.addEventListener("dragend", () => {
        item.classList.remove("dragging");
      });
      item.addEventListener("dragover", (event) => {
        event.preventDefault();
        item.classList.add("drag-over");
        if (event.dataTransfer) {
          event.dataTransfer.dropEffect = "move";
        }
      });
      item.addEventListener("dragleave", () => {
        item.classList.remove("drag-over");
      });
      item.addEventListener("drop", (event) => {
        event.preventDefault();
        item.classList.remove("drag-over");
        const sourceIndex = Number(event.dataTransfer?.getData("text/plain"));
        if (Number.isNaN(sourceIndex)) return;
        if (sourceIndex === index) return;
        movePageToIndex(sourceIndex, index, page.section || "body");
      });

      if (page.section === "front") {
        elements.outlineFront.appendChild(item);
      } else if (page.section === "back") {
        elements.outlineBack.appendChild(item);
      } else {
        elements.outlineBody.appendChild(item);
      }
    });
  }

  function getActiveParagraphLocation() {
    if (!activeParagraphId) return null;
    return findParagraphLocation(activeParagraphId);
  }

  function setPropertiesTab(tab) {
    if (!elements.propertiesElementsSection || !elements.propertiesFormattingSection) return;
    const showElements = tab === "elements";
    elements.propertiesElementsSection.classList.toggle("active", showElements);
    elements.propertiesFormattingSection.classList.toggle("active", !showElements);
    if (elements.propertiesTabElements) {
      elements.propertiesTabElements.classList.toggle("active", showElements);
    }
    if (elements.propertiesTabFormatting) {
      elements.propertiesTabFormatting.classList.toggle("active", !showElements);
    }
  }

  function updatePropertiesPanel() {
    if (!elements.propertiesCurrent) return;
    const location = getActiveParagraphLocation();
    const styleTag = location?.paragraph?.styleTag || "";
    elements.propertiesCurrent.textContent = styleTag ? styleLabel(styleTag) : "No selection";
    if (elements.propertiesClearBtn) {
      elements.propertiesClearBtn.disabled = !location;
    }

    document.querySelectorAll(".properties-item").forEach((button) => {
      const active = styleTag && button.dataset.style === styleTag;
      button.classList.toggle("active", active);
    });

    const format = location ? getParagraphFormat(location.paragraph) : getDocumentFormat();
    if (elements.formatFontFamily) {
      elements.formatFontFamily.value = format.fontFamily || "Times New Roman";
    }
    if (elements.formatFontSize) {
      elements.formatFontSize.value = format.fontSize;
    }
    if (elements.formatLineSpacing) {
      elements.formatLineSpacing.value = format.lineSpacing;
    }
    if (elements.formatAlign) {
      elements.formatAlign.value = format.textAlign || "left";
    }
    if (elements.formatIndent) {
      elements.formatIndent.value = format.indentCm;
    }
    if (elements.formatSpacingBefore) {
      elements.formatSpacingBefore.value = format.spacingBeforePt;
    }
    if (elements.formatSpacingAfter) {
      elements.formatSpacingAfter.value = format.spacingAfterPt;
    }
  }

  function applyStyleToActive(styleTag) {
    const location = getActiveParagraphLocation();
    if (!location) {
      alert(getContent("alerts.noParagraphSelected", "Select a paragraph to apply an element."));
      return;
    }
    location.paragraph.styleTag = styleTag;
    scheduleSave();
    renderCurrentPage();
  }

  function getDocumentFormat() {
    const settings = appState.docSettings;
    return {
      fontFamily: settings.fontFamily,
      fontSize: settings.fontSize,
      lineSpacing: settings.lineSpacing,
      textAlign: settings.textAlign || "left",
      indentCm: settings.firstLineIndentCm,
      spacingBeforePt: settings.spacingBeforePt,
      spacingAfterPt: settings.spacingAfterPt,
      indentEnabled: settings.indentEnabled,
    };
  }

  function pickNumber(value, fallback) {
    return typeof value === "number" && !Number.isNaN(value) ? value : fallback;
  }

  function getParagraphFormat(para) {
    const base = getDocumentFormat();
    const overrides = para?.format && typeof para.format === "object" ? para.format : {};
    return {
      fontFamily: overrides.fontFamily || base.fontFamily,
      fontSize: pickNumber(overrides.fontSize, base.fontSize),
      lineSpacing: pickNumber(overrides.lineSpacing, base.lineSpacing),
      textAlign: overrides.textAlign || base.textAlign,
      indentCm: pickNumber(overrides.indentCm, base.indentCm),
      spacingBeforePt: pickNumber(overrides.spacingBeforePt, base.spacingBeforePt),
      spacingAfterPt: pickNumber(overrides.spacingAfterPt, base.spacingAfterPt),
      indentEnabled:
        typeof overrides.indentEnabled === "boolean" ? overrides.indentEnabled : base.indentEnabled,
    };
  }

  function applyParagraphFormatting(partial) {
    const location = getActiveParagraphLocation();
    if (!location) {
      alert(getContent("alerts.noParagraphSelected", "Select a paragraph to apply formatting."));
      return;
    }
    if (!location.paragraph.format || typeof location.paragraph.format !== "object") {
      location.paragraph.format = {};
    }
    location.paragraph.format = { ...location.paragraph.format, ...partial };
    scheduleSave();
    renderCurrentPage();
    updatePropertiesPanel();
  }

  function getParagraphDocxSettings(para) {
    const base = appState.docSettings;
    const format = getParagraphFormat(para);
    return {
      ...base,
      fontFamily: format.fontFamily,
      fontSize: format.fontSize,
      lineSpacing: format.lineSpacing,
      textAlign: format.textAlign,
      spacingBeforePt: format.spacingBeforePt,
      spacingAfterPt: format.spacingAfterPt,
      firstLineIndentCm: format.indentCm,
      indentEnabled: format.indentEnabled,
    };
  }
  function renderCurrentPageStatsOnly() {
    const page = currentPage();
    if (!page) return;
    const stats = calcPageStats(page);
    elements.pageStats.textContent = `${stats.words} words | ${stats.chars} chars`;
  }

  function autoSizeTextarea(textarea) {
    if (!textarea) return;
    textarea.style.height = "auto";
    textarea.style.height = `${textarea.scrollHeight}px`;
  }

  function applyEditorParagraphStyles(textarea, config = {}) {
    if (!textarea) return;
    if (config.fontFamily) {
      textarea.style.fontFamily = config.fontFamily;
    }
    if (typeof config.fontSizePx === "number") {
      textarea.style.fontSize = `${config.fontSizePx}px`;
    }
    if (config.textAlign) {
      textarea.style.textAlign = config.textAlign;
    }
    if (typeof config.lineHeight !== "undefined") {
      textarea.style.lineHeight = String(config.lineHeight);
    }
    if (typeof config.indentPx === "number") {
      textarea.style.textIndent = config.indentPx ? `${config.indentPx}px` : "0";
    }
  }

  function render() {
    renderFileMeta();
    if (appState.ui.view === "setup") {
      renderSetup();
      return;
    }
    if (elements.setupScreen) {
      elements.setupScreen.classList.add("hidden");
    }
    if (elements.editorShell) {
      elements.editorShell.classList.remove("hidden");
    }
    renderToolbar();
    renderCurrentPage();
    renderOutline();
    updatePropertiesPanel();
  }

  const TRIM_SIZES = {
    "5x8": { width: 5, height: 8 },
    "5.25x8": { width: 5.25, height: 8 },
    "5.5x8.5": { width: 5.5, height: 8.5 },
    "6x9": { width: 6, height: 9 },
    "8.5x11": { width: 8.5, height: 11 },
  };

  function getTrimSize(trimSize) {
    return TRIM_SIZES[trimSize] || TRIM_SIZES["6x9"];
  }

  function inchesToPx(inches, dpi = 96) {
    return Math.round(inches * dpi);
  }

  function cmToPx(cm, dpi = 96) {
    return Math.round((cm / 2.54) * dpi);
  }

  function ptToPx(pt) {
    return Math.round(pt * 1.333);
  }

  function styleLabel(styleTag) {
    if (!styleTag) return "Normal";
    return STYLE_LABEL_MAP[styleTag] || styleTag;
  }

  function styleForExport(styleTag) {
    if (!styleTag) return "Normal";
    return STYLE_EXPORT_MAP[styleTag] || styleTag;
  }

  function resolvePageSide(pageSide, pageNumber) {
    if (pageSide === "left" || pageSide === "right") {
      return pageSide;
    }
    return pageNumber % 2 === 0 ? "left" : "right";
  }

  function applyTokens(text, pageNumber) {
    if (!text) return "";
    const title = baseFileName();
    return text
      .replace(/\{page\}/gi, String(pageNumber))
      .replace(/\{title\}/gi, title || "Manuscript");
  }

  function buildRunningBlock(text, pageNumberText, position) {
    const block = document.createElement("div");
    block.className = "book-page-running";

    const left = document.createElement("span");
    left.className = "running-left";
    left.textContent = position === "left" ? pageNumberText : text;

    const center = document.createElement("span");
    center.className = "running-center";
    center.textContent = position === "center" ? pageNumberText : "";

    const right = document.createElement("span");
    right.className = "running-right";
    if (position === "right") {
      right.textContent = pageNumberText;
    } else if (position !== "left") {
      right.textContent = "";
    }

    if (position === "left") {
      left.textContent = pageNumberText;
      right.textContent = text;
    }

    block.appendChild(left);
    block.appendChild(center);
    block.appendChild(right);
    return block;
  }

  function buildPageShell(settings, pageNumber, hostWidth, extraClass = "") {
    const trim = getTrimSize(settings.trimSize);
    const dpi = 96;
    const pageWidthPx = inchesToPx(trim.width, dpi);
    const pageHeightPx = inchesToPx(trim.height, dpi);
    const gutterPx = inchesToPx(settings.gutterIn || 0, dpi);
    const topPx = inchesToPx(settings.marginTopIn || 0.75, dpi);
    const bottomPx = inchesToPx(settings.marginBottomIn || 0.75, dpi);
    const insidePx = inchesToPx(settings.marginInsideIn || 0.75, dpi) + gutterPx;
    const outsidePx = inchesToPx(settings.marginOutsideIn || 0.5, dpi);
    const mirror = !!settings.mirrorMargins;
    const side = resolvePageSide(settings.pageSide, pageNumber);
    const isLeftPage = mirror && side === "left";
    const leftPx = isLeftPage ? outsidePx : insidePx;
    const rightPx = isLeftPage ? insidePx : outsidePx;

    const safeHost = Math.max(hostWidth || pageWidthPx, 320);
    const scale = Math.min(1, (safeHost - 24) / pageWidthPx);

    const frame = document.createElement("div");
    frame.className = `page-frame ${extraClass}`.trim();
    frame.style.width = `${Math.round(pageWidthPx * scale)}px`;
    frame.style.height = `${Math.round(pageHeightPx * scale)}px`;

    const page = document.createElement("div");
    page.className = `book-page ${extraClass}`.trim();
    page.style.fontFamily = settings.fontFamily;
    page.style.fontSize = `${Math.round(settings.fontSize * 1.333)}px`;
    page.style.lineHeight = settings.lineSpacing;
    page.style.width = `${pageWidthPx}px`;
    page.style.height = `${pageHeightPx}px`;
    page.style.padding = `${topPx}px ${rightPx}px ${bottomPx}px ${leftPx}px`;
    page.style.boxSizing = "border-box";
    page.style.transform = `scale(${scale})`;
    page.style.transformOrigin = "top left";

    const body = document.createElement("div");
    body.className = "book-page-body";

    const headerText = applyTokens(settings.headerText, pageNumber);
    const footerText = applyTokens(settings.footerText, pageNumber);
    const pageNumberPosition = settings.pageNumberPosition || "none";
    const showHeaderNumber = pageNumberPosition.startsWith("header");
    const showFooterNumber = pageNumberPosition.startsWith("footer");
    const headerPos = showHeaderNumber ? pageNumberPosition.split("-")[1] : null;
    const footerPos = showFooterNumber ? pageNumberPosition.split("-")[1] : null;

    const header =
      headerText || showHeaderNumber
        ? buildRunningBlock(headerText, showHeaderNumber ? String(pageNumber) : "", headerPos)
        : null;
    const footer =
      footerText || showFooterNumber
        ? buildRunningBlock(footerText, showFooterNumber ? String(pageNumber) : "", footerPos)
        : null;

    if (header) {
      header.classList.add("running-header");
      page.appendChild(header);
    }
    page.appendChild(body);
    if (footer) {
      footer.classList.add("running-footer");
      page.appendChild(footer);
    }

    frame.appendChild(page);
    return { frame, page, body, scale };
  }

  function movePage(offset) {
    const nextIndex = appState.currentPageIndex + offset;
    if (nextIndex < 0 || nextIndex >= appState.pages.length) return;
    appState.currentPageIndex = nextIndex;
    activeParagraphId = null;
    scheduleSave();
    render();
  }

  function exportTxt() {
    if (!validateBeforeExport()) return;
    const pages = appState.pages || [];
    if (!pages.length) {
      alert(getContent("alerts.noPagesToExport", "No pages to export."));
      return;
    }
    const text = pages
      .map((page) => page.paragraphs.map((p) => p.editedText).join("\n\n"))
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
    if (!validateBeforeExport()) return;
    if (!window.docx) {
      alert(getContent("alerts.docxMissing", "docx.js is missing. Ensure docx.umd.js is loaded."));
      return;
    }
    const pages = appState.pages || [];
    if (!pages.length) {
      alert(getContent("alerts.noPagesToExport", "No pages to export."));
      return;
    }

    const children = [];
    pages.forEach((page) => {
      page.paragraphs.forEach((para) => {
        children.push(buildDocxParagraph(para));
      });
    });

    const doc = new docx.Document({
      sections: [
        {
          properties: {
            page: {
              margin: buildDocxMargins(),
              size: buildDocxPageSize(),
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
    const settings = appState.docSettings;
    const top = inchesToTwip(settings.marginTopIn || 0.75);
    const bottom = inchesToTwip(settings.marginBottomIn || 0.75);
    const inside = inchesToTwip(settings.marginInsideIn || 0.75);
    const outside = inchesToTwip(settings.marginOutsideIn || 0.5);
    const gutter = inchesToTwip(settings.gutterIn || 0);
    return {
      top,
      bottom,
      left: inside + gutter,
      right: outside,
    };
  }

  function buildDocxPageSize() {
    const trim = getTrimSize(appState.docSettings.trimSize);
    return {
      width: inchesToTwip(trim.width),
      height: inchesToTwip(trim.height),
    };
  }

  function buildDocxParagraph(para) {
    const settings = getParagraphDocxSettings(para);
    const baseSize = settings.fontSize * 2;
    const styleTag = styleForExport(para.styleTag);
    const alignmentMap = {
      left: docx.AlignmentType.LEFT,
      center: docx.AlignmentType.CENTER,
      right: docx.AlignmentType.RIGHT,
      justify: docx.AlignmentType.JUSTIFIED,
    };
    const alignment = alignmentMap[settings.textAlign];
    const spacing = {
      before: ptToTwip(settings.spacingBeforePt),
      after: ptToTwip(settings.spacingAfterPt),
      line: Math.round(settings.lineSpacing * 240),
      lineRule: docx.LineRuleType.AUTO,
    };

    const indent = settings.indentEnabled
      ? { firstLine: cmToTwip(settings.firstLineIndentCm) }
      : {};

    const text = styleTag === "Poem" ? para.editedText : para.editedText.replace(/\n+/g, " ");

    if (styleTag === "Title") {
      return new docx.Paragraph({
        children: buildTextRuns(text, styleTag, settings.fontFamily, baseSize + 6),
        alignment: docx.AlignmentType.CENTER,
        spacing: { ...spacing, after: ptToTwip(12) },
      });
    }

    if (styleTag === "Heading1") {
      return new docx.Paragraph({
        children: buildTextRuns(text, styleTag, settings.fontFamily, baseSize + 2, { bold: true }),
        alignment,
        spacing: { ...spacing, before: ptToTwip(8), after: ptToTwip(6) },
      });
    }

    if (styleTag === "Quote") {
      return new docx.Paragraph({
        children: buildTextRuns(text, styleTag, settings.fontFamily, baseSize, { italics: true }),
        alignment,
        indent: { left: cmToTwip(0.8) },
        spacing,
      });
    }

    if (styleTag === "Poem") {
      return new docx.Paragraph({
        children: buildTextRuns(text, styleTag, settings.fontFamily, baseSize),
        alignment,
        spacing: { ...spacing, before: ptToTwip(2), after: ptToTwip(4) },
      });
    }

    return new docx.Paragraph({
      children: buildTextRuns(text, styleTag, settings.fontFamily, baseSize),
      alignment,
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

  function inchesToTwip(inches) {
    return Math.round(inches * 1440);
  }

  function validateBeforeExport() {
    if ((appState.project?.type || "print") !== "print") return true;
    const pageCount = appState.pages.length;
    const issues = validatePrintSettings(appState.docSettings, pageCount, appState.project.bleed);
    if (issues.length) {
      alert(`Cannot export for print:\n${issues.join("\n")}`);
      return false;
    }
    return true;
  }
  function bindEvents() {
    [
      elements.projectTitle,
      elements.projectAuthor,
      elements.projectLanguage,
      elements.projectType,
      elements.projectTrimSize,
      elements.projectBleed,
      elements.projectMarginTop,
      elements.projectMarginBottom,
      elements.projectMarginInside,
      elements.projectMarginOutside,
      elements.projectGutter,
      elements.projectMirror,
      elements.projectPageSide,
      elements.projectHeaderText,
      elements.projectFooterText,
      elements.projectPageNumberPosition,
      elements.projectFontFamily,
      elements.projectFontSize,
      elements.projectLineSpacing,
      elements.projectIndentSize,
      elements.projectSpaceBefore,
      elements.projectSpaceAfter,
    ].forEach((input) => {
      if (!input) return;
      input.addEventListener("change", () => {
        applyProjectSettingsFromSetup();
        renderSetup();
      });
    });

    if (elements.projectFileInput) {
      elements.projectFileInput.addEventListener("change", (event) => {
        const file = event.target.files[0];
        if (file) {
          pendingStoredState = null;
          if (elements.resumeBanner) {
            elements.resumeBanner.classList.add("hidden");
          }
          handleFileImport(file);
        }
        elements.projectFileInput.value = "";
      });
    }

    if (elements.projectStartBtn) {
      elements.projectStartBtn.addEventListener("click", startProject);
    }

    if (elements.projectSettingsBtn) {
      elements.projectSettingsBtn.addEventListener("click", openProjectSettings);
    }

    if (elements.paginateBtn) {
      elements.paginateBtn.addEventListener("click", () => {
        const confirmText = getContent(
          "alerts.repaginateConfirm",
          "Repaginating will rebuild pages and reset edits. Continue?"
        );
        if (confirm(confirmText)) {
          const text = collectEditedText();
          repaginateFromText(text);
          scheduleSave();
          render();
        }
      });
    }

    if (elements.loadSessionBtn && elements.sessionInput) {
      elements.loadSessionBtn.addEventListener("click", () => {
        elements.sessionInput.click();
      });
    }

    if (elements.sessionInput) {
      elements.sessionInput.addEventListener("change", (event) => {
        const file = event.target.files[0];
        if (file) {
          pendingStoredState = null;
          if (elements.resumeBanner) {
            elements.resumeBanner.classList.add("hidden");
          }
          loadSession(file);
        }
        elements.sessionInput.value = "";
      });
    }

    if (elements.exportSessionBtn) elements.exportSessionBtn.addEventListener("click", exportSession);
    if (elements.exportTxtBtn) elements.exportTxtBtn.addEventListener("click", exportTxt);
    if (elements.exportDocxBtn) elements.exportDocxBtn.addEventListener("click", exportDocx);

    if (elements.docSettingsBtn && elements.docSettingsModal) {
      elements.docSettingsBtn.addEventListener("click", () => {
        elements.docSettingsModal.classList.remove("hidden");
      });
    }

    if (elements.docSettingsClose && elements.docSettingsModal) {
      elements.docSettingsClose.addEventListener("click", () => {
        elements.docSettingsModal.classList.add("hidden");
      });
    }

    if (elements.addFrontBtn) elements.addFrontBtn.addEventListener("click", () => addBlankPage("front"));
    if (elements.addBodyBtn) elements.addBodyBtn.addEventListener("click", () => addBlankPage("body"));
    if (elements.addBackBtn) elements.addBackBtn.addEventListener("click", () => addBlankPage("back"));

    if (elements.propertiesTabElements) {
      elements.propertiesTabElements.addEventListener("click", () => setPropertiesTab("elements"));
    }
    if (elements.propertiesTabFormatting) {
      elements.propertiesTabFormatting.addEventListener("click", () => setPropertiesTab("formatting"));
    }
    if (elements.propertiesClearBtn) {
      elements.propertiesClearBtn.addEventListener("click", () => applyStyleToActive("Normal"));
    }
    document.querySelectorAll(".properties-item").forEach((button) => {
      button.addEventListener("click", () => {
        const tag = button.dataset.style;
        if (!tag) return;
        applyStyleToActive(tag);
      });
    });

    if (elements.formatFontFamily) {
      elements.formatFontFamily.addEventListener("change", (event) => {
        applyParagraphFormatting({ fontFamily: event.target.value });
      });
    }
    if (elements.formatFontSize) {
      elements.formatFontSize.addEventListener("change", (event) => {
        const value = Number(event.target.value);
        if (Number.isNaN(value)) return;
        applyParagraphFormatting({ fontSize: value });
      });
    }
    if (elements.formatLineSpacing) {
      elements.formatLineSpacing.addEventListener("change", (event) => {
        const value = Number(event.target.value);
        if (Number.isNaN(value)) return;
        applyParagraphFormatting({ lineSpacing: value });
      });
    }
    if (elements.formatAlign) {
      elements.formatAlign.addEventListener("change", (event) => {
        applyParagraphFormatting({ textAlign: event.target.value });
      });
    }
    if (elements.formatIndent) {
      elements.formatIndent.addEventListener("change", (event) => {
        const value = Number(event.target.value);
        if (Number.isNaN(value)) return;
        applyParagraphFormatting({ indentCm: value });
      });
    }
    if (elements.formatSpacingBefore) {
      elements.formatSpacingBefore.addEventListener("change", (event) => {
        const value = Number(event.target.value);
        if (Number.isNaN(value)) return;
        applyParagraphFormatting({ spacingBeforePt: value });
      });
    }
    if (elements.formatSpacingAfter) {
      elements.formatSpacingAfter.addEventListener("change", (event) => {
        const value = Number(event.target.value);
        if (Number.isNaN(value)) return;
        applyParagraphFormatting({ spacingAfterPt: value });
      });
    }

    if (elements.prevPageBtn) elements.prevPageBtn.addEventListener("click", () => movePage(-1));
    if (elements.nextPageBtn) elements.nextPageBtn.addEventListener("click", () => movePage(1));

    if (elements.pageSection) {
      elements.pageSection.addEventListener("change", (event) => {
        const page = currentPage();
        if (!page) return;
        page.section = event.target.value;
        scheduleSave();
        renderOutline();
      });
    }

    if (elements.pageTitle) {
      elements.pageTitle.addEventListener("input", (event) => {
        const page = currentPage();
        if (!page) return;
        page.title = event.target.value;
        scheduleSave();
        renderOutline();
      });
    }

    if (elements.resumeBtn) {
      elements.resumeBtn.addEventListener("click", () => {
        if (!pendingStoredState) return;
        appState = hydrateState(pendingStoredState);
        pendingStoredState = null;
        elements.resumeBanner.classList.add("hidden");
        render();
      });
    }

    if (elements.dismissBtn) {
      elements.dismissBtn.addEventListener("click", () => {
        pendingStoredState = null;
        localStorage.removeItem(STORAGE_KEY);
        elements.resumeBanner.classList.add("hidden");
      });
    }

    [
      elements.docTrimSize,
      elements.docPageSide,
      elements.docHeaderText,
      elements.docFooterText,
      elements.docPageNumberPosition,
      elements.docFontFamily,
      elements.docFontSize,
      elements.docLineSpacing,
      elements.docMargins,
      elements.docMarginTop,
      elements.docMarginBottom,
      elements.docMarginInside,
      elements.docMarginOutside,
      elements.docGutter,
      elements.docMirrorMargins,
      elements.docIndentEnabled,
      elements.docIndentSize,
      elements.docSpaceBefore,
      elements.docSpaceAfter,
    ].forEach((input) => {
      if (!input) return;
      input.addEventListener("change", () => {
        if (elements.docTrimSize) {
          appState.docSettings.trimSize = elements.docTrimSize.value;
        }
        if (elements.docPageSide) {
          appState.docSettings.pageSide = elements.docPageSide.value;
        }
        if (elements.docHeaderText) {
          appState.docSettings.headerText = elements.docHeaderText.value;
        }
        if (elements.docFooterText) {
          appState.docSettings.footerText = elements.docFooterText.value;
        }
        if (elements.docPageNumberPosition) {
          appState.docSettings.pageNumberPosition = elements.docPageNumberPosition.value;
        }
        if (elements.docFontFamily) {
          appState.docSettings.fontFamily = elements.docFontFamily.value;
        }
        if (elements.docFontSize) {
          appState.docSettings.fontSize = Number(elements.docFontSize.value);
        }
        if (elements.docLineSpacing) {
          appState.docSettings.lineSpacing = Number(elements.docLineSpacing.value);
        }
        if (elements.docMargins) {
          appState.docSettings.marginsPreset = elements.docMargins.value;
        }
        if (elements.docMarginTop) {
          appState.docSettings.marginTopIn = Number(elements.docMarginTop.value);
        }
        if (elements.docMarginBottom) {
          appState.docSettings.marginBottomIn = Number(elements.docMarginBottom.value);
        }
        if (elements.docMarginInside) {
          appState.docSettings.marginInsideIn = Number(elements.docMarginInside.value);
        }
        if (elements.docMarginOutside) {
          appState.docSettings.marginOutsideIn = Number(elements.docMarginOutside.value);
        }
        if (elements.docGutter) {
          appState.docSettings.gutterIn = Number(elements.docGutter.value);
        }
        if (elements.docMirrorMargins) {
          appState.docSettings.mirrorMargins = elements.docMirrorMargins.checked;
        }
        if (elements.docIndentEnabled) {
          appState.docSettings.indentEnabled = elements.docIndentEnabled.checked;
        }
        if (elements.docIndentSize) {
          appState.docSettings.firstLineIndentCm = Number(elements.docIndentSize.value);
        }
        if (elements.docSpaceBefore) {
          appState.docSettings.spacingBeforePt = Number(elements.docSpaceBefore.value);
        }
        if (elements.docSpaceAfter) {
          appState.docSettings.spacingAfterPt = Number(elements.docSpaceAfter.value);
        }
        scheduleSave();
        renderCurrentPage();
      });
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





