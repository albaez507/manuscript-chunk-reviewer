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
  buttons: {
    prevPage: "Prev Page",
    nextPage: "Next Page",
    docSettingsClose: "Close",
  },
  options: {
    docFontFamily: {
      times: "Times New Roman",
      georgia: "Georgia",
      garamond: "Garamond",
    },
    pageSide: {
      auto: "Auto (by page #)",
      right: "Right page",
      left: "Left page",
    },
    pageNumber: {
      none: "Off",
      headerRight: "Header right",
      headerCenter: "Header center",
      footerRight: "Footer right",
      footerCenter: "Footer center",
    },
    trimSize: {
      "5x8": "5\" x 8\"",
      "5.25x8": "5.25\" x 8\"",
      "5.5x8.5": "5.5\" x 8.5\"",
      "6x9": "6\" x 9\"",
      "8.5x11": "8.5\" x 11\"",
    },
    margins: {
      clean: "Clean Book",
    },
  },
  panels: {
    docSettingsTitle: "Document Settings",
    docSettingsSubtitle: "Clean template export (not a round-trip of original styles).",
  },
  docSettings: {
    trimSize: "Trim size",
    pageSide: "Page side",
    headerText: "Header text",
    footerText: "Footer text",
    pageNumberPosition: "Page number position",
    fontFamily: "Font family",
    baseFontSize: "Base font size (pt)",
    lineSpacing: "Line spacing",
    marginsPreset: "Margins preset",
    marginTop: "Top margin (in)",
    marginBottom: "Bottom margin (in)",
    marginInside: "Inside margin (in)",
    marginOutside: "Outside margin (in)",
    gutter: "Gutter (in)",
    mirrorMargins: "Mirror margins",
    firstLineIndent: "First-line indent (cm)",
    spacingBefore: "Spacing before (pt)",
    spacingAfter: "Spacing after (pt)",
  },
  notes: {
    docSettingsNote: "Title, Heading, Quote, and Poem styles are applied at export time based on each paragraph's style tag.",
  },
  labels: {
    pageLabel: "Page",
  },
  messages: {
    noPages: "No pages loaded",
    emptyState: "Import a file to begin page layout.",
    storedSession: "Stored session",
  },
  alerts: {
    missingMammoth: "Mammoth library is missing. Ensure mammoth.browser.min.js is loaded.",
    docxMissing: "docx.js is missing. Ensure docx.umd.js is loaded.",
    docxParseError: "Unable to parse .docx file.",
    emptyFile: "The file is empty.",
    docxFallbackError: "Unable to read this .docx. Try saving it again as a Word .docx file.",
    loadSessionError: "Unable to load session JSON.",
    noPagesToExport: "No pages to export.",
    repaginateConfirm: "Repaginating will rebuild pages and reset edits. Continue?",
    deletePageConfirm: "Delete this page?",
    noParagraphSelected: "Select a paragraph to apply an element.",
  },
};
