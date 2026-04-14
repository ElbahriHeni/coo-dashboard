# COO CRM React Dashboard

A downloadable React + Vite dashboard for COO KPI monitoring.

## Features
- KPI cards: quotations, policies converted, conversion rate, expected GWP for Motor, actual GWP
- Filters: month, line of business, region, source, department
- Charts for pipeline, quotation status, department workload, conversion by LOB, policies by duration, expected GWP by phase, actual GWP by region
- Detailed data table
- Mock data included in `src/data/dashboardData.js`

## Run locally

```bash
npm install
npm run dev
```

Then open the local URL shown in the terminal.

## Build

```bash
npm run build
npm run preview
```

## Customize
Replace the mock dataset in `src/data/dashboardData.js` with your CRM data or hook the app to your API.
