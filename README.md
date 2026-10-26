# AutoFMScrapper

Pulls a company's annual financials and builds an Excel workbook with a three-statement model and a DCF. The history is filled in for you. The forecast and valuation inputs are yours to enter.

[Add one line here on why you built it.]

## About this project

I built AutoFMScrapper myself as part of my finance and Python learning. I have also used AI tools to some extent to structure the code and fix bugs, and I've managed the code and tested the output. Mistakes are still possible, so check the numbers against the original filings before relying on them.

## Disclaimer

This tool is for learning and research only. It is not financial advice. The data comes from third-party sources and may be incomplete or wrong, and every forecast and valuation depends on the inputs you enter. Use it at your own risk.

## Quick start

1. Run `Setup AutoFMScrapper.cmd` once. It creates the Python environment and installs `requirements.txt`. You need Python installed on Windows.
2. Double-click `Open AutoFMScrapper.cmd`.
3. Enter a ticker and market, then click **Create my workbook**.

Each run gets its own folder and ZIP in `Output/`, so old workbooks never get overwritten. You can reuse a saved snapshot or fetch fresh data. For Indian companies you can load a Screener annual Excel/CSV export instead.

There's an AAPL example under Recent workbooks if you want to try the blank-input model first.

## Using the workbook

1. **History:** Green numbers came from the source. Yellow blanks mean the source didn't have that figure, so check the filing and fill it in. Only enter 0 if it really is 0.
2. **Assumptions:** Fill the five yellow forecast columns (type `5%` for 5%). Use `0` for a year with no borrowing, dividends or share financing. To reuse the same inputs for all five years, fill F9:F23, select F9:J23 and press Ctrl+R.
3. **Statements:** The income statement, balance sheet and cash flow run off your inputs. Read the checks under them. Negative cash means you need funding. I don't hide it with an automatic loan.
4. **DCF inputs:** Back on Assumptions, add your discount rate, long-run growth (must be below the discount rate), share count, other claims and non-operating assets. Share price is optional.
5. **DCF:** Uses cash flow from the same forecast. If an input is missing, that result stays blank. Blank doesn't mean worthless.

Everything is plain Excel formulas, so it recalculates as you type. You don't need to rerun the app.

## What's in the workbook

- **Statements:** three historical years plus five forecast years
- **History:** every normalized year available, with editable source corrections
- **Raw statements:** all original line items and periods from the source
- **Sources:** the original download/import, normalized history, field map, input cell map and the generation check report

How much history you get depends on the data provider. Nothing is made up. Missing items stay blank.

## How the model works

Revenue follows your growth input. Gross margin sets cost of sales, and EBITDA margin sets cash operating costs. Depreciation and interest are based on opening fixed assets and opening debt. Tax is charged on positive profit only, with no loss carryforwards. Receivables, inventory and payables come from your days assumptions, and other balances stay flat. Equity rolls forward with profit, dividends and share financing (buybacks count as negative share financing).

Not included: acquisitions, disposals, stock comp, FX forecasts, and detailed debt or tax schedules. I kept it simple so the statements stay linked and the inputs stay readable. It won't give you a company-specific forecast. That part is yours.

The DCF discounts year-end free cash flow to the firm at WACC, adds cash and non-operating assets, and subtracts debt and other claims. Terminal value grows from year five at your rate, so check that your investment assumptions support that growth. Share count is in millions.

Reference: [CFA Institute, free cash flow valuation](https://www.cfainstitute.org/insights/professional-learning/refresher-readings/2026/free-cash-flow-valuation)

This is for non-financial companies. Banks and insurers need a different structure.

## Running from source

Command line:

```
.\.venv\Scripts\python.exe AutoFMScrapper.py --ticker AAPL
```

Add `--refresh` to fetch new data, `--import-file path.xlsx` for a Screener export, or `--output new-folder` to pick the output folder.

Workbook creation uses Node with `@oai/artifact-tool`. It's bundled on my machine but not in this repo. On another computer, set `FM_NODE` to your Node path and `FM_ARTIFACT_MODULES` to a `node_modules` folder containing that package. Checking the calculations natively needs Microsoft Excel on Windows. Once saved, the workbook has no links back to Python.

## Code layout

- `autofmscrapper/gui.py`: desktop interface
- `autofmscrapper/runner.py`: source files, workbook generation, saved runs
- `autofmscrapper/model.py`: worksheet layout and Excel formulas
- `autofmscrapper/data.py`: source mapping
- `write.mjs`: formatting and export
- `fetch_data.py`: Yahoo downloads

For linting and formatting, install `requirements-dev.txt` and run:

```
.\.venv\Scripts\python.exe -m ruff check autofmscrapper AutoFMScrapper.py fetch_data.py
```

Use `ruff format` on the same paths. Rules are in `pyproject.toml`.
