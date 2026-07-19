# portfolio_tracker
📊 A real-time stock portfolio tracker in Python. Fetches live market valuations via the Yahoo Finance API (yfinance) and automatically generates a highly stylized, reactive corporate Excel dashboard (openpyxl) featuring dynamic formula injection and asset allocation tracking.

# Real-Time Stock Portfolio Tracker

A lightweight Python tool that fetches **live stock prices** via `yfinance`
and generates a formatted **Excel** portfolio summary — market value per
holding, allocation percentage, and a grand total — using live spreadsheet
formulas rather than hardcoded numbers.

## Features

- Fetches real-time price and company name for each ticker in your portfolio
- Calculates market value (`quantity × price`) and allocation % per holding
- Exports a styled `.xlsx` report with a dated filename
- Built as a class (`PortfolioTracker`) so it's easy to extend or import
- Includes basic error handling for tickers that fail to fetch

## Requirements

- Python 3.8+
- `yfinance`
- `openpyxl`

## Installation

```bash
pip install yfinance openpyxl
```

## Usage

1. Open the script and edit `PORTFOLIO_DATA` with your own tickers and share
   quantities:

   ```python
   PORTFOLIO_DATA = {
       "AAPL": 25,
       "MSFT": 15,
       "GOOGL": 10,
       "AMZN": 12,
       "NVDA": 8,
   }
   ```

2. Run the script:

   ```bash
   python portfolio_tracker.py
   ```

3. Find the generated report in the `output/` folder:

   ```
   output/portfolio_2026-07-19.xlsx
   ```

## How It Works

| Step | Method | Description |
|---|---|---|
| 1 | `fetch_live_prices()` | Pulls current price and company name for each ticker via `yfinance`, using `fast_info` first and falling back to `info` if needed |
| 2 | `generate_excel_report()` | Writes ticker, company, quantity, and price to a sheet, then adds live formulas for market value, allocation %, and the portfolio total |

The Excel formulas (e.g. `=C2*D2` for market value, `=E2/$E$8` for
allocation) recalculate automatically in Excel if you open the file and
change a quantity — the numbers aren't frozen at export time.

## Output Columns

| Column | Description |
|---|---|
| Ticker | Stock symbol |
| Company | Company short name |
| Quantity | Shares held (from your input) |
| Price | Live price at time of fetch |
| Market Value | `Quantity × Price` |
| Allocation % | Share of total portfolio value |

## Project Structure

```
.
├── portfolio_tracker.py
├── output/
│   └── portfolio_<date>.xlsx   # generated on each run
└── README.md
```

## Notes & Limitations

- Requires an internet connection — prices are fetched live from Yahoo
  Finance at runtime.
- If a ticker fails to fetch (e.g. rate limiting, invalid symbol, or a
  temporary API issue), it's skipped and a warning is printed rather than
  crashing the whole run.
- Prices reflect the market state at the moment the script runs; re-run it
  to refresh.
- This is a portfolio *valuation* tool, not investment advice — it doesn't
  make buy/sell recommendations.

## Possible Extensions

- Add day-over-day price change columns
- Log each run to a CSV for historical trend tracking
- Add a pie chart visualizing allocation
- Support multiple currencies
