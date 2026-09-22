# GT Player Dashboard

A data analytics and visualization system for tracking individual player performance across a 96-member competitive gaming alliance. Built with Excel/VBA (backend) and GitHub Pages (frontend), with a live data pipeline connecting the two.

🔗 **[Live Dashboard](https://kimpossible7544.github.io/Player-Dashboard/)**

---

## What It Does

Provides individual player analytics including score trends, rank history, goal performance, and resting week breakdowns. Data is processed in an Excel workbook via VBA macros and surfaced in a web dashboard that updates whenever the workbook is refreshed.

### Features

- **Per-player analytics dashboard** — score trends, weekly rank, missed goals, resting score, and resting rank
- **Active vs. Resting week segmentation** — metrics automatically split between competitive weeks and designated rest periods
- **Multi-week trend charts** — line charts for score and rank trends, bar charts for missed goals, with resting weeks labeled inline
- **Roster management** — 96-player roster with ID-based lookup, alias normalization, and duplicate detection
- **Arena Power and HQ Power tracking** — historical power tracking with baseline comparison and growth-since-joining metrics
- **Role-based access** — roster-ID login system with master password for full alliance view

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Excel / VBA (`.xlsm`) |
| Data pipeline | Dropbox CDN + SheetJS |
| Frontend | HTML, CSS, JavaScript (vanilla) |
| Deployment | GitHub Pages |
| Automation | Power Automate |

---

## System Architecture

```
Excel Workbook (.xlsm)
├── Roster sheet (96 players, ID mapping, aliases)
├── Weekly score sheets (daily scores, goal tracking)
├── Player Tracking sheet (aggregated stats, ranks)
├── Arena Power sheet (power history, growth tracking)
├── Score Import sheet (daily import workflow)
└── Week Settings sheet (active/resting week flags)
        │
        ▼
    VBA Macro Engine
    ├── ImportScoresFromScoreImport
    ├── UpdatePlayerTrackingFromRosterAndWeeks
    ├── BuildPlayerDashboard / RefreshDashboard
    └── PopulateWeeklyIDs (roster ID normalization)
        │
        ▼
    Dropbox CDN
        │
        ▼
    GitHub Pages Dashboard
    └── gt_data.js (SheetJS parser + data pipeline)
```

---

## VBA Macro System

The workbook contains a full automation engine handling:

- **Score import** — pulls name/score pairs from a staging sheet, normalizes names against the roster, and writes to the correct weekly sheet column
- **Player tracking** — aggregates weekly totals, calculates overall and active-week ranks, counts goal misses (active weeks only), and computes resting scores
- **Dashboard refresh** — builds interactive player cards with chart objects (score trend, missed goals, rank trend) styled with a dark command-center aesthetic
- **ID population** — resolves player names to roster IDs using normalized name matching, alias lookup, and Levenshtein distance fallback
- **Name normalization** — handles Unicode, Cyrillic, Greek, and special character variants across in-game display names
- **WPX Arena block** — `RefreshWPXArenaBlock` rebuilds the WPX-start lifetime comparison and WPX history block at column 220+ on `Arena Power`; run `ImportWPXData` first so the hidden `WPX_ArenaPower` copy is available

---

## Key Design Decisions

**Why Excel as a backend?**  
The alliance management workflow happens in Excel. Using it as the data store eliminates a sync step and keeps the system maintainable by non-developers.

**Why no server?**  
GitHub Pages is free and zero-maintenance. Dropbox CDN serves the workbook file directly to the browser, where SheetJS parses it client-side. No API keys, no hosting costs, no infrastructure.

**Why resting week segmentation?**  
Competitive alliances designate low-activity weeks where scoring goals don't apply. Penalizing players for those weeks would distort performance metrics — so the system separates active and resting data at every level: ranks, averages, goal counts, and chart labels.

---

## Related

See the [GT](https://github.com/Kimpossible7544/WPX) repo for the score tracker and Tyrant event tracking tools.
