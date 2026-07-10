---
name: testing-gt-dashboard
description: Test the GT Player Dashboard pages (Standings, Dashboard, Projections) end-to-end. Use when verifying UI changes, data ordering, or chart rendering.
---

# Testing GT Player Dashboard

## Overview
The GT Player Dashboard is a static site with 3 pages that fetch and render data from a Dropbox-hosted Excel file (`GTFinal5.3.26.xlsm`).

## Environment Setup

1. Clone the repo and check out the branch to test
2. Start a local HTTP server (needed for CORS when fetching from Dropbox):
   ```bash
   cd /path/to/Player-Dashboard
   python3 -m http.server 8080 &
   ```
3. Pages are available at:
   - `http://localhost:8080/wpx_standings.html` — Player Standings
   - `http://localhost:8080/wpx_dashboard.html` — Player Dashboard (Scoreboard)
   - `http://localhost:8080/wpx_projections.html` — Score Projections

## Data Flow

- All 3 pages load `gt_data.js` which fetches the Excel file from Dropbox at runtime
- The Excel "Player Tracking" sheet contains weekly score data
- `gt_data.js` parses week columns starting at column 6, each week taking 2 columns (Total + Rank)
- Week labels, total columns, and rank columns are extracted into arrays that all pages consume
- The Excel file stores weeks newest-first (new weeks are added to the left in Excel)

## Key Testing Areas

### Week Ordering
- Weeks should display chronologically: oldest on left, newest on right
- **Standings**: Check the week column headers in the table (scroll right to see all)
- **Dashboard**: Check x-axis labels on Weekly Score Trend, Weekly Rank, and Daily Score charts
- **Projections**: Check Score History + Projection chart x-axis AND the "Week by Week Breakdown" table
- If ordering is broken, the newest week would appear as the first/leftmost column or label

### Player Selection
- Dashboard and Projections have a player dropdown selector
- The first player in the dropdown is selected by default on page load
- Test with different players to verify data renders correctly

### Data Loading
- Check "Data loaded" timestamp appears below the page title (confirms successful fetch)
- If data fails to load, an error message appears instead
- Check browser console for JavaScript errors during page load

## Tips

- The Dropbox URL may change if the user re-uploads the file. If data fails to load, check `gt_data.js` for the current URL.
- The pages use Chart.js for rendering charts — x-axis labels are derived from the `weekLabels` array in `gt_data.js`
- No authentication is required; this is a fully static site
- All pages share the same data source (`gt_data.js`), so a fix in the data layer affects all pages simultaneously
