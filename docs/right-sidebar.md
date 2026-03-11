# Right Sidebar (Kindle Create Style) - Feb 28, 2026

## Scope Implemented (Option A)
- Replaced the old Review Queue panel with a Kindle Create–style **Properties** panel.
- Two tabs: **Elements** and **Formatting**.
- Current Element display with **Clear** action.
- Elements grouped per Kindle Create:
  - Title Pages
  - Book Start and End Pages
  - Common Elements
  - Book Body
- Formatting controls:
  - Font family / size / line spacing
  - Paragraph alignment
  - First-line indent
  - Spacing before/after
  - (Drop cap placeholder)

## Behavior
- Applies Element styles to the **active paragraph**.
- Active paragraph auto-selects the first paragraph if none is selected.
- Formatting controls update the document settings and re-render the page.
- Editor text reflects alignment/indent/spacing visually.
- DOCX export maps extended Kindle Create styles to base styles:
  - Book Title -> Title
  - Chapter Title -> Heading1
  - Block Quote -> Quote
  - Poem -> Poem
  - Others -> Normal (or closest)

## Files Updated
- `index.html`
  - Replaced right panel markup with Properties panel and tabs.
- `styles.css`
  - New styles for tabs, groups, items, fields, and active states.
- `app.js`
  - New UI bindings for properties panel.
  - Style mapping + formatting application.
  - Helper functions for style labels and editor paragraph styling.

## Notes / Open Items
- Drop cap is currently a placeholder.
- Style application is paragraph-level (matches Kindle Create behavior).
- If you want InDesign-like controls, we can extend formatting and add style presets.

## Next Steps (when we continue)
1. Visual tuning: more Kindle Create / InDesign polish (spacing, typographic scale, icons).
2. Add style presets (e.g., Chapter Title, Quote) to influence formatting defaults.
3. Decide if formatting stays in right panel or moves to a bottom dock (Grammarly-like).
