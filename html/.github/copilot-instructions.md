## Purpose
This repo is a minimal static HTML project (single page) used as a classroom reference clock. These instructions help AI coding agents be immediately productive editing and testing the site.

## Quick repo summary
- Single primary file: `clock.html` (root of this workspace). It contains inline CSS and JavaScript and is the only UI artifact.
- Workspace file: `html.code-workspace` (VS Code workspace settings may exist here).
- No build system, package manifest, or tests found in this workspace.

## What to change and where (concrete pointers)
- Visual/style: edit the `<style>` block near the top of `clock.html`.
- Text content shown in the UI: inside the `<section class="card">` and the `#editable` element (a contenteditable div) at the top of the page.
- Clock logic and formatting: the inline `<script>` at the bottom of `clock.html`. Key DOM IDs:
  - `hoursMinutes`, `seconds`, `ampm` — render the time
  - `dateDisplay` — friendly date text
  - `tzDisplay` — timezone label
  - `toggle24` — toggles 24/12-hour display

## Important patterns and conventions (project-specific)
- Single-file SPA: styles and logic are inline; prefer small, local changes rather than splitting into many files unless adding a clear need.
- Accessibility hints are used: `aria-live="polite"` on the time display and `role="region"` on the card. Preserve these attributes when refactoring markup.
- Time rendering uses the browser Intl APIs (no external timezone libs). Keep dependency-free changes unless you add a clear, documented reason.

## Developer workflows (how to run and debug locally)
- No build step — open `clock.html` directly in a browser for simple edits.
- Recommended local server options (from the repo root) — choose one:

PowerShell (if Python is installed):
```powershell
python -m http.server 8000
# then open http://localhost:8000/clock.html
```

Node (if using npm and http-server):
```powershell
npx http-server -p 8000
# then open http://localhost:8000/clock.html
```

- VS Code + Live Server extension: right-click `clock.html` -> "Open with Live Server". This is the fastest edit-refresh loop for designers.

## Tests, CI, and external integration
- No tests or CI discovered. There are no external API integrations; the page relies on browser Intl and system clock only.

## Good-first-change examples (explicit tasks an AI can do now)
- Change the accent color: update `--accent` in `:root` and adjust `.time` styles.
- Add a 12/24 preference stored in localStorage: augment the existing toggle handler and read/write `use24Hour`.
- Improve accessibility: add `aria-label` attributes to the toggle button and ensure color contrast meets WCAG where possible.

## When to ask for human input
- Don't change authentication or production deployment (no CI exists here) without confirmation.
- If you propose adding packages (npm, pip), document why and what commands to run; ask for permission before adding dependencies.

## Files to open first (order of investigation)
1. `clock.html` — primary source of truth (styles, markup, JS)
2. `html.code-workspace` — workspace preferences

If you want, I can also: add a minimal README, wire a tiny test harness, or split the JS/CSS into separate files and add a simple local-start script. Tell me which direction you prefer.
