# AutoFMScrapper

A separate three-statement model and DCF. It downloads the company's annual
history and leaves the forecast and valuation inputs for you to enter.

Double-click **Open AutoFMScrapper.cmd**. Enter a ticker and market, then click
**Create my workbook**. You can reuse a saved source snapshot or request a new
download. For an Indian company, you can select a Screener annual Excel/CSV
export. Each run gets its own folder and ZIP inside Output; earlier files are
preserved.

An AAPL example is already in Recent workbooks. Select its row and click
Open workbook to try the blank-input model before fetching another company.

## First time in the workbook

1. **Check History.** Downloaded numbers are green. Yellow blanks mean the
   source did not provide that figure. If a required value is missing, check
   the filing and enter it here. Do not use zero unless it really is zero.
2. **Open Assumptions.** Fill the five yellow forecast columns. For 5%, type
   `5%`. Enter `0` for a year with no borrowing, dividends or share financing.
   The last-actual column is a reference, not a recommended forecast.
   To keep the same inputs for all five years, fill F9:F23, select F9:J23
   and press Ctrl+R. You can then change individual years.
3. **Read Statements.** The income statement, balance sheet and cash flow use
   those inputs. Fill one year at a time if you prefer. Look at the checks
   below the statements. Negative cash says funding is needed; it is not
   hidden by an automatic loan.
4. **Enter the DCF inputs.** Add your discount rate, long-run growth, share
   count, other claims and non-operating assets on Assumptions. Enter zero
   for claims/assets you have checked and found absent. Growth must be below
   the discount rate. A current share price is optional.
5. **Read DCF.** It takes cash flow from the same forecast. Missing inputs
   leave the related result blank, so a blank value does not mean worthless.

All forecast and DCF input cells start empty, including zero-value assumptions.
The formulas are already linked. These formulas calculate in Excel when you
enter your own inputs; there is no need to rerun the Python app after editing.

## What is included

- Three visible historical years and five forecast years on Statements.
- Every available normalized year on History, with editable source corrections.
- All original source line items and available periods on Raw statements.
- The original downloaded JSON/import, normalized history, source field map
  and input cell map in Sources, plus the generation check report.

The provider controls how much history is available. The model never invents
older years, missing items or analyst assumptions. Period labels and reporting
basis are as supplied. Review source gaps in the app and the original filing.
This format is for non-financial companies; banks and insurers need a different
statement and valuation structure.

## How the simple model works

Revenue follows your growth input. Gross margin sets cost of sales; EBITDA
margin sets cash operating costs. Depreciation and interest use opening fixed
assets and opening debt. Tax is charged on positive profit, without tax-loss
carryforwards. Receivables, inventory and payables follow your day assumptions.
Other asset and liability balances stay constant. Equity rolls forward with
profit, dividends and cash share financing. Share buybacks are negative share
financing; the DCF share count is a separate input you must review.

There are no acquisitions, disposals, stock-compensation schedules, separate FX
forecasts or advanced debt/tax modules. Historical cash-flow reconciliation
items remain visible. These simplifications keep the three statements linked
and the inputs readable; they do not supply a company-specific forecast.

DCF uses year-end free cash flow to the firm, discounted at WACC. It adds cash
and eligible non-operating assets, then deducts debt and other claims. Shares
are in millions; the per-share formula converts the workbook's monetary units
to currency units. Terminal cash flow grows from year five at your entered
rate; check that investment supports that growth.

[DCF method reference — CFA Institute](https://www.cfainstitute.org/insights/professional-learning/refresher-readings/2026/free-cash-flow-valuation).

## Local setup

This folder has its own code, Python environment, cache, history and outputs.
It does not import Studio's model engine or modify its workbooks. The launcher
uses the environment inside this folder. If it is missing, the setup script
creates it and installs requirements.txt. Python must be installed on Windows.

Run **Setup AutoFMScrapper.cmd** to repair that Python environment. Workbook creation also
uses the bundled Node/Artifact Tool runtime already on this computer. On another
computer, set `FM_NODE` to Node and `FM_ARTIFACT_MODULES` to a node_modules folder
containing `@oai/artifact-tool`. Do not modify the bundled dependency folder.
Native calculation verification needs Microsoft Excel on Windows. Once saved,
the workbook has no links back to Python and can be used on its own in Excel.

CLI: `.\.venv\Scripts\python.exe AutoFMScrapper.py --ticker AAPL`. Add `--refresh` for a new fetch,
`--import-file path.xlsx` for a Screener export, or `--output new-folder`.

## Working on the source

The desktop interface is in `autofmscrapper/gui.py`. `runner.py` handles source files,
workbook generation and saved runs. `model.py` defines the worksheet layout and
Excel formulas; `write.mjs` applies formatting and exports the workbook.
Source mapping is in `data.py`, and Yahoo downloads are in `fetch_data.py`.

Install `requirements-dev.txt` for the Python formatter and linter. Formatting
rules are in `pyproject.toml`. From this folder, run
`.\.venv\Scripts\python.exe -m ruff check autofmscrapper AutoFMScrapper.py fetch_data.py`
and use `ruff format` with the same paths to format an edit.
