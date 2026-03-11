# Book Layout Studio (working name)

Current internal name in code/UI is still `Manuscript Chunk Reviewer`.

## Product Direction
This app is being repositioned as a lightweight replacement for:
- Kindle Create (book-focused author workflow)
- Adobe InDesign (basic-to-mid layout freedom, not complex pro desktop publishing)

Target outcome: help an indie author or book formatter produce KDP Print-ready book interiors faster, with less friction than full desktop tools.

## Ideal User
- Indie author
- Book formatter working on non-complex to medium-complex projects

## Core Jobs To Be Done
1. Set up a book project with proper KDP print settings (trim, margins, bleed, pagination behavior).
2. Edit text and apply per-paragraph style/format controls in a Kindle Create-like flow.
3. Build front matter, body, and back matter with visible page layout.
4. Export a final file suitable for Amazon KDP Print upload.

## Scope Decision (MVP)
Primary focus now:
- Layout and pagination
- Practical text editing + formatting

Not in MVP now:
- AI grammar/spelling layer
- Advanced InDesign-level automation

AI-based correction is planned as a later layer after layout/export fundamentals are stable.

## Language Support
The product should support:
- Spanish
- English
- Dutch

UI/content must stay localization-friendly (`content.js` as source of labels/messages).

## Export Direction
Target final export: KDP Print-compatible PDF.

Current implementation status:
- Implemented: `.txt`, `.docx`, JSON session
- Planned next: PDF export path aligned with KDP Print requirements

## Why "Chunk" Is No Longer Central
Originally, "chunk" represented reviewing text in smaller pieces instead of a large Word page.
Now that the app renders real book trim sizes (for example `6 x 9`), the workflow is page/layout-first.
So product language should move from "chunk review" to "book layout and formatting."

## Current Architecture (High Level)
- Single-page browser app (`index.html`, `app.js`, `styles.css`, `content.js`)
- Offline-first session state via `localStorage` (`mcr_state_v1`)
- Page-based model with sectioned outline:
  - Front matter
  - Body
  - Back matter
- Paragraph-level element and formatting controls in right-side Properties panel
- Import: `.txt`, `.md`, `.docx`

## Near-Term Roadmap
1. Rebrand naming across UI/docs away from "chunk reviewer".
2. Stabilize KDP print workflow (project setup, pagination, page preview fidelity).
3. Implement PDF export pipeline for KDP Print upload.
4. Add QA pass for multilingual UI (ES/EN/NL labels and validation messages).
5. Add optional grammar/spelling correction layer (post-MVP).

## Repository Structure
```
/index.html
/styles.css
/app.js
/content.js
/assets/
/docs/
```
