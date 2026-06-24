# Fleet Dashboard — Setup & Usage Guide

## Folder Structure

```
fleet-dashboard/
│
├── build.js              ← run this every month to build dashboards
├── package.json          ← dependencies (run npm install once)
│
├── template/
│   └── dashboard.html    ← the dashboard template (do not edit)
│
├── clients/
│   ├── menengai/
│   │   ├── Summary_Menengai.xlsx          ← replace each month
│   │   ├── Trip_Menengai_May-2026.xlsx    ← keep all, add new ones
│   │   └── Trip_Menengai_Jun-2026.xlsx
│   │
│   └── another-client/
│       ├── Summary_AnotherClient.xlsx
│       └── Trip_AnotherClient_May-2026.xlsx
│
└── output/
    ├── Menengai_Dashboard.html     ← send this to client
    └── Anotherclient_Dashboard.html
```

---

## File Naming Rules

| File Type        | Must start with | Example |
|-----------------|-----------------|---------|
| Monthly Summary | `Summary`       | `Summary_Menengai.xlsx` |
| Trip Detail     | `Trip`          | `Trip_Menengai_May-2026.xlsx` |

⚠️ Only ONE Summary file per client folder at a time.
   Delete the old one before dropping in the new one.

Trip files accumulate — keep all previous months in the folder.

---

## First-Time Setup

1. Install Node.js from https://nodejs.org (LTS version)
2. Open Command Prompt in this folder
3. Run: npm install
4. Done — you're ready to build

---

## Monthly Workflow

```
End of month
  ↓
1. Delete old Summary file from clients\menengai\
2. Drop new Summary file into clients\menengai\
3. Drop new Trip file into clients\menengai\
4. Open Command Prompt in this folder
5. Run: node build.js menengai
6. Find output at: output\Menengai_Dashboard.html
7. Email to client
```

---

## Adding a New Client

1. Create a new folder under clients\  
   Example: clients\new-client\

2. Drop their files in:
   - Summary_NewClient.xlsx
   - Trip_NewClient_May-2026.xlsx

3. Build:
   node build.js new-client

---

## Commands

```
node build.js menengai          ← build Menengai dashboard
node build.js another-client    ← build another client
```

---

## Dashboard Features

- **Live View** — full 12-month overview, all vehicles
- **Monthly Snapshots** — frozen view per month
- **Category Filter** — filter to any department/sub-client
- **Daily Heatmap** — vehicle × day grid for months with trip data
- **Vehicle Drill-down** — click any vehicle for full history
- **Vehicle Table** — searchable, sortable, idle vehicles flagged

---

## Troubleshooting

**"No Summary*.xlsx found"**  
→ Make sure your summary file starts with "Summary"

**"Template not found"**  
→ Make sure dashboard.html is in the template/ folder

**Dashboard shows wrong data**  
→ Check that old Summary file was deleted before adding new one

**Charts not showing**  
→ Open the HTML in Chrome or Edge (not Internet Explorer)
