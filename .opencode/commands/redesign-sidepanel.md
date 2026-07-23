---
description: Implement the Drag to Sheets side-panel redesign
---

Implement the side-panel UI refresh defined in `design.md`.

Before editing:

1. Read `design.md` completely.
2. Read `README.md`, `sidepanel/sidepanel.html`, `sidepanel/sidepanel.css`, and the relevant UI-rendering portions of `sidepanel/sidepanel.js`.
3. Inspect the tests and `package.json` so verification matches the repository.
4. Identify all DOM IDs and class names consumed by JavaScript.
5. Produce a concise implementation plan before changing files.

Implementation requirements:

- Make the real extension interface more polished, compact, and professional.
- Treat `design.md` as the source of truth.
- Preserve every existing workflow and feature.
- Preserve existing DOM IDs unless you update all corresponding JavaScript safely.
- Prefer changes to `sidepanel/sidepanel.html` and `sidepanel/sidepanel.css`.
- Change `sidepanel/sidepanel.js` only where required for safe state classes, accessibility, or revised UI structure.
- Do not alter parser, cleaner, merger, Google API, exporter, or storage behavior.
- Do not add a framework, remote assets, analytics, telemetry, or unnecessary dependencies.
- Continue using Lucide icons and the existing local assets.
- Keep the interface realistic for a Chrome side panel; do not reproduce a wide promotional mockup inside the extension.
- Support widths from 320px through 600px without page-level horizontal overflow.
- Ensure keyboard focus, semantic controls, accessible labels, disabled states, and reduced motion remain correct.

Work through the UI states methodically:

- Empty and drag-over states
- URL import collapsed and expanded
- One and multiple files loaded
- Separate and merge modes
- Smart mapping, mapping review, and custom mapping
- Cleaning options and duplicate sub-options
- Preview and long-table content
- Disabled, loading, success, warning, and error states
- Narrow and wide side-panel widths

Verification:

1. Run `npm test`.
2. Fix regressions caused by the redesign.
3. Review the changed files for accidental behavior changes.
4. Check for selectors that no longer match the HTML.
5. Confirm there is no page-level horizontal overflow.
6. Summarize the files changed, the visual improvements, functional safeguards, and test results.

Additional direction from the user, when provided:

$ARGUMENTS
