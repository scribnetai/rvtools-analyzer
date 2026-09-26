# Changelog

## 2026-09-26
- Added 💾 Projects: named saves in this browser, portable JSON export/import, and automatic session restore (your last analysis reloads on revisit). Large datasets fall back to Export when they exceed browser storage. Privacy FAQ updated: saves live only in your browser's localStorage, never uploaded.
- Downloadable HTML report v3: full VM inventory (all VMs, 11 columns), guest OS mix, VMware Tools status, hosts-per-cluster lists, uncapped reclaim/snapshot/right-size lists, print-friendly page breaks.
- Expanded landing page: 3-step How it works, tab-by-tab read table, honest limitations, worked licensing examples, FAQ grew 5→12.
- Fixed: stale report bug — browsers were caching js/app.js; added cache-busting (?v=N) to CSS/JS includes.
- Added: this changelog section, rendered from CHANGELOG.md.
