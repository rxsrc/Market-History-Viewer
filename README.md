# Market History Viewer

A desktop stock history viewer built with Python and Tkinter. Type a ticker, see its full daily open/close history in a color-coded table, a 3-month price chart, and the company name. Compare several stocks in split panes, search by date, and export to CSV, Excel or Google Sheets.

> **Coded by [Claude](https://www.anthropic.com/claude), an AI assistant made by Anthropic.** The project owner directed the design and feature set; Claude wrote the code.

---

## Features

| Feature | Details |
|---|---|
| **Company name lookup** | Enter `TTWO` and the title reads `TTWO - Take-Two Interactive Software, Inc.` |
| **Full history table** | Every trading day: Date, Open, Close, Net ($), Change (%). Newest first. Green rows for gains, red for losses. |
| **3-month chart** | Close-price chart in the top right of each pane, below the title and above the data table. Green if up over the period, red if down. |
| **Date search** | Jump straight to a specific day. If the market was closed, it goes to the nearest earlier trading day. |
| **Date range filter** | Show only rows between a From and To date (either can be left blank). |
| **Export to CSV** | Saves exactly what the pane is showing (including any filter). |
| **Export to Excel** | `.xlsx` with the same green/red row colors and proper number, date and percent formats. |
| **Export to Google Sheets** | One-click upload with a Google service account, or a clipboard-and-paste fallback with no setup. |
| **Multithreading** | History, company name and exports all run on background threads, so the UI never freezes. |
| **Split panes** | Up to 4 stocks at once, side by side or stacked, each with its own table and chart. |
| **Keyboard shortcuts** | Fast control without touching the mouse (see below). |

---

## Installation

Requires **Python 3.9+** (Tkinter ships with most Python installs; on some Linux distros install `python3-tk`).

```bash
git clone <your-repo-url>
cd <your-repo-folder>
pip install yfinance pandas matplotlib openpyxl gspread
python market_history_viewer.py
```

`gspread` is only needed for the Google Sheets upload.

---

## Usage

1. Type a ticker (for example `AAPL`, `TTWO`, `MSFT`) in the sidebar and press **Enter** or click **Load Preview**.
2. The right side appears with the company name, stats, 3-month chart and the full history table.
3. Use the sidebar to search by date, filter a range, export, or add panes.

The sidebar always controls the **active pane**, which is shown with a blue border. Click a pane to make it active.

### Searching by date

- **Find Date**: enter `YYYY-MM-DD` and press Enter. The row is highlighted and scrolled into view.
- **Filter date range**: fill in From and/or To, then **Apply**. **Clear** restores the full history.

### Split panes (compare stocks)

Press **Ctrl+N** to add a pane, then enter another ticker. Each pane keeps its own stock, filter and chart. **F6** switches between side-by-side and stacked layouts, and the dividers between panes can be dragged to resize.

---

## Keyboard shortcuts

| Shortcut | Action |
|---|---|
| `Enter` | Load ticker / jump to date |
| `Ctrl+N` | Split: add another stock pane |
| `Ctrl+W` | Close the active pane |
| `Ctrl+1` to `Ctrl+4` | Switch active pane |
| `Ctrl+L` | Focus the ticker box |
| `Ctrl+G` | Focus the "go to date" box |
| `F5` | Refresh the active pane (bypasses cache) |
| `F6` | Toggle side-by-side / stacked |
| `Ctrl+S` | Export CSV |
| `Ctrl+E` | Export Excel |
| `Ctrl+Shift+G` | Export to Google Sheets |
| `F1` | Show shortcut help |

---

## Exporting

### CSV and Excel
Choose a location in the save dialog. Both formats export the rows currently shown in the active pane, so a date-filtered view exports as filtered. The Excel file keeps the green/red styling, a frozen header row and filters.

### Google Sheets
There are two modes:

**1. One-click upload (recommended).** 
1. Create a Google Cloud project and enable the **Google Sheets API** and **Google Drive API**.
2. Create a **service account** and download its JSON key.
3. Save the key as `service_account.json` next to `market_history_viewer.py` (or set the `GOOGLE_SERVICE_ACCOUNT` environment variable to its path).
4. Press **Ctrl+Shift+G**, enter your Google account email, and the app creates the sheet, shares it with you, and opens it.

**2. No setup fallback.** Without a key file, the app copies the data to your clipboard and opens a blank Google Sheet. Click cell A1 and paste.

> Keep `service_account.json` private. Add it to your `.gitignore` so it is never committed.

---

## How it works

- **Data source:** [`yfinance`](https://github.com/ranaroussillon/yfinance), an unofficial library for Yahoo Finance data.
- **GUI:** Tkinter with `ttk.Treeview` for the table and `ttk.PanedWindow` for split panes.
- **Charts:** Matplotlib embedded in Tkinter (`FigureCanvasTkAgg`).
- **Threading:** a `ThreadPoolExecutor` runs network and file work. Workers never touch Tk; they post results to a queue that the main thread polls every 100 ms. Stale results (for example from a pane you closed or reloaded) are discarded.
- **Responsiveness:** the history table fills in chunks so very long histories (decades of data) don't block the window.
- **Caching:** loaded tickers are cached for 5 minutes. **F5** forces a fresh download.

---

## Project structure

```
market_history_viewer.py   # the entire app
README.md
.gitignore                 # recommended: service_account.json, __pycache__/
```

---

## Troubleshooting

- **"No data found"**: check the ticker symbol. Use Yahoo Finance symbols (for example `BRK-B`, not `BRK.B`).
- **Company name missing**: Yahoo sometimes doesn't return it. The ticker still loads and the title shows the symbol alone.
- **Layout toggle does nothing**: some older Tk versions can't change pane orientation after creation.
- **Google Sheets error**: make sure both the Sheets and Drive APIs are enabled and the key file path is correct.

---

## Disclaimer

Market data comes from Yahoo Finance via an unofficial library and may be delayed, incomplete or inaccurate. This tool is for informational purposes only and is **not financial advice**.

---

## Credits

Written by **Claude** (Anthropic) at the request of the project owner.

## License

Add the license of your choice (for example MIT) as a `LICENSE` file.
