/*
Centralized UI copy for Manuscript Chunk Reviewer.
This file is a single place to manage labels, prompts, and messages.
*/
window.MCR_CONTENT = {
  brand: {
    title: "Manuscript Chunk Reviewer",
    subtitleNoFile: "No file loaded",
  },
  banner: {
    prompt: "Resume previous session?",
    resume: "Resume",
    dismiss: "Dismiss",
  },
  toolbar: {
    import: "Import",
    session: "Session",
    export: "Export",
    docSettings: "Doc Settings",
    chunkSizeTarget: "Chunk size target",
    layout: "Layout",
    engine: "Engine",
    mode: "Mode",
    applyScope: "Apply to",
    actions: "Actions",
    compare: "Compare",
    compareToggle: "Side-by-side",
  },
  buttons: {
    analyze: "Analyze",
    suggest: "Suggest",
    rechunk: "Rechunk",
    reset: "Reset",
    skip: "Skip",
    approve: "Approve",
    undo: "Undo",
    redo: "Redo",
    exportTxt: "Final .txt",
    exportDocx: "Final .docx",
    loadSession: "Load .json",
    exportSession: "Export .json",
    docSettingsOpen: "Open",
    docSettingsClose: "Close",
  },
  options: {
    layout: {
      editor: "Editor Focus (70/30)",
      balanced: "Balanced (50/50)",
      review: "Review Focus (40/60)",
    },
    engine: {
      local: "Local (offline)",
      ai: "AI (placeholder)",
    },
    mode: {
      clean: "Clean Structure",
      punct: "Punctuation/Spelling",
      translation: "Faithful Translation (future)",
      style: "Style (future)",
    },
    applyScope: {
      focused: "Focused paragraph",
      all: "All in chunk",
    },
    docFontFamily: {
      times: "Times New Roman",
      georgia: "Georgia",
      garamond: "Garamond",
    },
    margins: {
      clean: "Clean Book",
    },
  },
  panels: {
    approvedHistoryTitle: "Approved History",
    approvedHistorySubtitle: "Inline diff + traceable change log",
    docSettingsTitle: "Document Settings",
    docSettingsSubtitle: "Clean template export (not a round-trip of original styles).",
  },
  docSettings: {
    fontFamily: "Font family",
    baseFontSize: "Base font size (pt)",
    lineSpacing: "Line spacing",
    marginsPreset: "Margins preset",
    firstLineIndent: "First-line indent (cm)",
    spacingBefore: "Spacing before (pt)",
    spacingAfter: "Spacing after (pt)",
  },
  notes: {
    docSettingsNote: "Title, Heading, Quote, and Poem styles are applied at export time based on each paragraph's style tag.",
  },
  labels: {
    chunkLabel: "Chunk",
    paragraphPrefix: "P",
  },
  messages: {
    noChunks: "No chunks loaded",
    emptyState: "Import a file to begin chunk review.",
    historyEmpty: "Approved chunks will appear here with diffs and trace logs.",
    noChangeLog: "No change log entries.",
    noFlags: "No flags.",
    storedSession: "Stored session",
  },
  alerts: {
    missingMammoth: "Mammoth library is missing. Ensure mammoth.browser.min.js is loaded.",
    docxMissing: "docx.js is missing. Ensure docx.umd.js is loaded.",
    docxParseError: "Unable to parse .docx file.",
    loadSessionError: "Unable to load session JSON.",
    noApprovedChunks: "No approved chunks to export.",
    rechunkConfirm: "Rechunking will reset approvals and edits. Continue?",
  },
  shortcuts: {
    approve: "Ctrl/Cmd + Enter",
    reset: "Ctrl/Cmd + Shift + R",
    undo: "Ctrl/Cmd + Z",
    redo: "Ctrl/Cmd + Shift + Z",
  },
};
