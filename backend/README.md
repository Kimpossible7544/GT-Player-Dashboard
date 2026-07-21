# GT Player Dashboard State API

A small FastAPI backend that stores per-week, per-player `Done` checkbox state for the rank-shift Promote/Demote tables in the GT Player Dashboard.

## Endpoints

- `GET /health` — health check
- `GET /state?week=<week label>` — get all checkbox states for a week
- `POST /state` — save a checkbox state (`{ week, player, status }`)
- `GET /all` — dump all stored states

All mutating requests require the header `X-API-Key: GT2026`.
