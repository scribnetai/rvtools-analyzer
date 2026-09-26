# RVTools Analyzer

A browser-based RVTools export analyzer for presales SEs. Drop an RVTools `.xlsx` export and get:

- **Executive summary** — environment totals + auto-generated findings in plain English
- **Cluster breakdowns** — hosts, cores, VMs, vCPU:pCore, memory overcommit, storage per cluster
- **Per-core licensing impact** — Broadcom's 16-core-minimum-per-socket math, phantom cores, and an interactive **refresh scenario modeler** (what if we consolidate onto denser hosts?)
- **Host inventory** — sortable table with license-core math per host
- **Storage** — datastore health, thin-vs-thick splits, top consumers
- **VM analysis** — searchable inventory, powered-off reclaim candidates, snapshot hygiene, right-size review shortlists
- **One-click HTML briefing report** — standalone file you can email to the team or customer

## Privacy: zero data retention, by design

- The file is read with the browser's `FileReader` API and parsed **in memory** by a vendored copy of SheetJS — there is no server and no upload endpoint.
- The app makes **zero network requests** with your data (all JS/CSS is local; it works offline after the page loads).
- Nothing is written to `localStorage`, `IndexedDB`, or cookies. "Clear session data" (or closing the tab) wipes everything.

## Run it

Hosted: `https://scribnetai.github.io/rvtools-analyzer/` — or open `index.html` directly, no build step.

No export handy? Click **"Try it with demo data"** for a realistic synthetic environment generated in your browser.

## How the licensing math works

Broadcom licenses vSphere per physical core with a **minimum of 16 cores per CPU socket**:

```
license cores per host = sockets × max(cores per socket, 16)
phantom cores          = license cores − physical cores
```

Phantom cores are licenses you pay for but can't use. The analyzer rolls this up per cluster and the scenario modeler shows what a refresh onto denser hosts would save.

> ⚠️ Indicative math for scoping conversations — not a quote. Always validate against an official Broadcom quote.

## Project layout

```
index.html          landing + dashboard shell
css/styles.css      dark theme
js/app.js           parser, analyzer, renderers, report generator (all client-side)
lib/xlsx.full.min.js  vendored SheetJS (offline parsing, no CDN dependency)
```

## Sister project

Part of the SE command-center family — the main dashboard links here: https://scribnetai.github.io/se-command-center/
