---
name: testing-gt-dashboard
description: Test the GT Player Dashboard pages (Standings, Dashboard, Projections) end-to-end. Use when verifying UI changes, login, data ordering, growth/power cards, or chart rendering.
---

# Testing GT Player Dashboard

## Overview
The GT Player Dashboard is a static site (HTML/CSS/JS + SheetJS) whose pages fetch and render
data from a Dropbox-hosted Excel workbook (`GTStatsFINAL.xlsm`). All data parsing lives in
`gt_data.js`; a fix there affects all pages.

## Environment Setup
1. Check out the branch to test.
2. Start a local HTTP server (needed for CORS when fetching from Dropbox):
   ```bash
   cd /path/to/GT-Player-Dashboard
   python3 -m http.server 8080 &
   ```
3. Pages:
   - `http://localhost:8080/gt_standings.html` — Standings
   - `http://localhost:8080/gt_dashboard.html` — Player Dashboard (has login)
   - `http://localhost:8080/gt_projections.html` — Projections
4. Hard-refresh (Ctrl+Shift+R) after switching branches so you don't get a cached page/workbook.

## Login (gt_dashboard.html)
The dashboard is gated by a login overlay (`doLogin()`):
- **Master password** `GT2026` → full alliance view, player dropdown VISIBLE.
- **Individual login** = a player's numeric **roster ID** → dashboard locks to that player,
  dropdown HIDDEN. IDs come from the Roster sheet (`idToPlayer`).
- Invalid entry → "Invalid ID. Please try again.", overlay stays.
- To find a valid roster ID without asking the user, download the workbook and read the Roster
  sheet: IDs in columns A/E/I/M (0,4,8,12), names in B/F/J/N (1,5,9,13). Known: `1030` =
  Kimpossible7544.
- After logout, re-login must restore `contentArea` (a past bug left it blank —
  `unlockDashboard()` must set `contentArea.style.display = ''`). Worth a quick regression check.

## Data Flow
- Pages load `gt_data.js`, which fetches the workbook at runtime (cache-busted with
  `&cb=<timestamp>` + `cache:no-store` so updates show without a manual hard refresh).
- "Player Tracking" sheet: `Weekly Total (label)` / `Weekly Rank (label)` column pairs.
- GT filters weeks to `SEASON_START` (currently July 13, 2026) — older columns exist in the
  file but are hidden, so the dashboard may show only the current week.
- Weekly Total is a precomputed cell; if it looks stale vs the daily sheet, the workbook needs
  a recalc/save (dashboard just reads whatever the cell holds).

## Growth / Power card (Growth Since Joining)
- Rendered by `renderGrowthCard(p)` only if `p.growth` exists; sourced from the "Arena Power"
  sheet (Arena Power + HQ/"personal" power, with Then→Now + delta and a baseline date badge).
- **Cross-team merge**: for roster IDs in `CROSS_TEAM_PLAYER_IDS` (players on both the previous
  WPX team and GT, matched by stable roster ID), `gt_data.js` also fetches `WPXStatsFinal.xlsm`
  and merges the WPX Arena/HQ power history so growth spans both teams. GT's Arena Power sheet
  may be empty, so a cross-team player's "Then" values / baseline date can only come from WPX —
  a good adversarial check (values impossible from GT alone prove the merge ran).
- Scope check: a cross-team player NOT in `CROSS_TEAM_PLAYER_IDS` (e.g. Addis, ID 1056) should
  show NO growth card. Only ~4 IDs overlap between the rosters (1007, 1030, 1056, 1087).

## Key Testing Areas
- **Week ordering**: oldest on left, newest on right (charts x-axis; Standings headers;
  Projections breakdown table).
- **Player selection**: master view dropdown defaults to first player; switch players to verify.
- **Data loading**: "Data loaded" timestamp appears under the title on success; errors render a
  message. Check the console for `[GT]` logs (roster IDs loaded, cross-team merge, week count).

## Tips
- The Dropbox URLs may change if the user re-uploads. If data fails to load, check the
  `DROPBOX_URL` / `WPX_DROPBOX_URL` constants in `gt_data.js`.
- Matching across teams is by **roster ID** (stable), not name — similar names with different
  IDs are different people.
- For ad-hoc verification of expected values, download the workbook and parse with
  `openpyxl` (`pip install openpyxl`) rather than eyeballing the sheet.

## Devin Secrets Needed
- None. The site is public/static and the Dropbox links are unauthenticated ("anyone with the
  link"). The master password `GT2026` and roster IDs are read from the workbook, not secrets.
