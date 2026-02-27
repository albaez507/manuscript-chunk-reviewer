# Style Guide

## Typography
- Display: `Fraunces` (used for the brand title).
- Body/UI: `Work Sans` (used for controls, paragraphs, and labels).
- Default body size: 14px in editor textareas.
- UI label size: 10–12px, uppercase tracking for control labels.

## Color Tokens
Defined in `styles.css` under `:root`:
- `--bg`: #f6f1e7 (page background)
- `--panel`: #ffffff (panel surface)
- `--panel-contrast`: #fdf9f2 (panel header)
- `--ink`: #1f1b16 (primary text)
- `--muted`: #6f665d (secondary text)
- `--accent`: #2a6f5e (primary action)
- `--accent-2`: #d9b46b (secondary accent)
- `--border`: #e2d8c7 (borders)
- `--shadow`: rgba(31, 27, 22, 0.08) (soft shadow)
- `--diff-add`: #d9f3e4 (diff additions)
- `--diff-del`: #f7d6d2 (diff deletions)
- `--diff-neutral`: #f4efe7 (diff neutral)

## Layout
- Toolbar: sticky with a warm gradient background and soft shadow.
- Workspace: two-column split driven by `--split-left` / `--split-right`.
- Responsive: collapses to a single column under 1100px.

## Components
- Buttons: rounded pill, hover lift, primary in `--accent`.
- Panels: card surfaces with `--radius` and soft shadow.
- Paragraph blocks: nested cards with textarea editing.
- Banners: floating resume prompt, pill shape.
- Modals: centered card with overlay scrim.
- Diff: inline highlights for additions/deletions/neutral.
- Flags & logs: compact pill/box patterns with muted text.
- Empty state: centered illustration plus message.

## Spacing & Radius
- Base radius: 14px (`--radius`).
- Internal padding: 12–20px for cards and toolbars.
- Grid gaps: 12–16px across sections.
