# DES Comparator

A single-file web app for comparing two AIR catalog exports (Excel) and tracking role assignment changes across products.

## Features

- **File comparison** — upload two `.xlsx` exports and instantly see what changed: new entries, people added/removed, and role changes
- **Filters** — filter results by role, product ID, EID, cell, certification status, and change type
- **Email integration** — send templated notifications per person via a local Python server (Outlook) or mailto fallback
- **Role Distribution chart** — visual histogram of how many people are assigned per product for any selected role, with drill-down modals showing the full list per product
- **Monthly summary** — KPI snapshot exportable as PNG/PDF for sharing
- **Dark mode** — toggleable theme

## Usage

Open `DES_Comparator_v4.4.html` directly in a browser — no installation or build step required.

1. Upload the **previous** AIR export (File 1) and the **current** one (File 2)
2. Click **Compare**
3. Use the filters to drill into specific roles, cells, or change types
4. Click **📊 Distribution** to open the role distribution panel

## Email server (optional)

For sending emails via Outlook, run the local Flask server:

```bash
cd server
pip install flask pywin32
python server.py
```

The app will detect the server at `http://localhost:5000` and use it automatically. If the server is not running, it falls back to mailto.

## Tech stack

- Vanilla HTML / CSS / JavaScript (single file, no framework)
- [SheetJS](https://sheetjs.com/) for Excel parsing
- [Chart.js](https://www.chartjs.org/) + chartjs-plugin-datalabels for the distribution chart
- [html2canvas](https://html2canvas.hertzen.com/) for screenshot export
- Python + Flask + pywin32 for the optional email server
