# Changelog

## 2026-09-28
- Added a floating Feedback button (bottom-right) that opens a dialog to send feedback via email — topic chips, optional name, and message, addressed to the site owner with the app name in the subject.

## 2026-09-28
- Fixed: refresh-scenario slider track fill was stuck at 50% — the theme paints the track with a `--fill` CSS variable, but nothing set it. Sliders now update their fill live as you drag (same fix as server-sizer).
## 2026-09-28
- TLS certificate provisioned for the `rvtools-analyzer.scribnet.io` custom domain (GitHub's stuck DNS check was reset 2026-09-28); HTTPS is now enforced on the site. App-switcher menu links switched from legacy `scribnetai.github.io` URLs to direct `https://<app>.scribnet.io` URLs for all 10 apps (footer/launcher links updated likewise). This entry also covers the net-zero CNAME delete/re-add commits from the DNS-check reset, which carried no changelog entries. Touched: index.html, js/app-switcher.js.


## 2026-09-26
- Physgun-style vibe reskin: Outfit typeface, blue→cyan gradient headline words, eyebrow labels + blue glyph tiles on section headings, scroll-reveal on landing sections, pill-style tabs, and rounder cards/panels with a faint blue glow. Refresh scenario modeler now uses glowing sliders (step 1) with live pill value badges, min/mid/max scale labels, and real-time rebuilds — plus an animated today-vs-scenario license-core bar comparison and big gradient key numbers. All dashboard bars (cluster utilization, phantom-core waste, thin/thick stacked meter, reclaim + snapshot candidates) animate in on render via data-w widths. FAQ items got emoji prefixes. No logic changes — parsing, math, Projects, and privacy guarantees untouched.
- Added 💾 Projects: named saves in this browser, portable JSON export/import, and automatic session restore (your last analysis reloads on revisit). Large datasets fall back to Export when they exceed browser storage. Privacy FAQ updated: saves live only in your browser's localStorage, never uploaded.
- Downloadable HTML report v3: full VM inventory (all VMs, 11 columns), guest OS mix, VMware Tools status, hosts-per-cluster lists, uncapped reclaim/snapshot/right-size lists, print-friendly page breaks.
- Expanded landing page: 3-step How it works, tab-by-tab read table, honest limitations, worked licensing examples, FAQ grew 5→12.
- Fixed: stale report bug — browsers were caching js/app.js; added cache-busting (?v=N) to CSS/JS includes.
- Added: this changelog section, rendered from CHANGELOG.md.

## 2026-09-27
- Added top-left app-switcher dropdown on the brand mark: one-click jumps to every app in the suite (full index, this page marked).
