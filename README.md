# Manuscript Chunk Reviewer

## Overview
Manuscript Chunk Reviewer is a lightweight, offline-first browser app for reviewing long manuscripts in manageable chunks. It imports `.txt`, `.md`, or `.docx` files, splits them into paragraph chunks, lets you edit with optional analysis/suggestions, and exports the approved result to `.txt` or `.docx` with clean formatting.

## Core Workflow
1. Import a manuscript file.
2. Review and edit paragraphs chunk-by-chunk.
3. Analyze or suggest edits with the local rules engine.
4. Approve or skip each chunk.
5. Export final text or a formatted `.docx`.

## Key Features
- Chunk-based review with adjustable target size.
- Side-by-side compare view (original vs edited).
- Local rules engine for basic analysis and suggestions.
- Per-paragraph undo/redo (Ctrl/Cmd + Z, Ctrl/Cmd + Shift + Z).
- Staging area to send paragraphs to the right panel as modular cards.
- Preview toggle to view staged paragraphs as a book page.
- Session persistence in `localStorage` plus JSON import/export.
- Approved history with inline diff and change log.
- Custom document settings for `.docx` export.
- Centralized UI copy in `content.js` for easy updates.

## Data Model (High Level)
- `sourceParagraphs`: raw paragraphs from the imported file.
- `chunks[]`: chunk objects with paragraph arrays.
- Each paragraph tracks `originalText`, `editedText`, `styleTag`, `changeLog`, `flags`, and `history`.
- `docSettings`: export settings applied during `.docx` creation.

## External Dependencies
- `mammoth.browser.min.js` for `.docx` import.
- `diff_match_patch.js` for inline diff rendering.
- `docx.umd.js` for `.docx` export.

All are loaded via CDN in `index.html`.

## Storage & Sessions
- Auto-saves state to `localStorage` under key `mcr_state_v1`.
- “Resume previous session” banner appears when saved state exists.
- JSON session files can be exported/imported for portability.

## Important Behaviors
- Rechunking resets approvals and edits because it rebuilds from source paragraphs.
- “Apply to” can target the focused paragraph or all in the chunk.
- Clean-mode suggestions may split long sentences and reduce connector chaining.
- Exported `.docx` is a clean template, not a round-trip of original styles.

## Limitations / Notes
- AI engine is a placeholder (no API integration yet).
- No global undo stack beyond per-paragraph history.

## Folder Structure
```
/index.html
/styles.css
/app.js
/content.js
/assets/
  /bg/
  /icons/
  /illustrations/
  /stickers/
  /placeholders/
/docs/
  STYLE_GUIDE.md
  ASSET_LIST.md
```
