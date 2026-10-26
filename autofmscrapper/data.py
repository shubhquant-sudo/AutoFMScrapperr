"""Source normalization and immutable snapshot cache. No forecast calculations."""

import base64
import csv
import hashlib
import io
import json
import math
import re
from datetime import date, datetime, timezone
from pathlib import Path

BASE = Path(__file__).resolve().parents[1]
ROOT = BASE
FIELDS = {
    "revenue": ("Revenue", "income", ["TotalRevenue"]),
    "cogs": ("Cost of sales", "income", ["CostOfRevenue"]),
    "gross_profit": ("Gross profit", "income", ["GrossProfit"]),
    "ebitda": ("EBITDA (provider definition)", "income", ["EBITDA"]),
    "da": (
        "Depreciation & amortization",
        "income",
        ["DepreciationAndAmortizationInIncomeStatement", "ReconciledDepreciation"],
    ),
    "ebit": ("Operating profit (EBIT)", "income", ["OperatingIncome"]),
    "interest": (
        "Interest expense",
        "income",
        ["InterestExpense", "InterestExpenseNonOperating"],
    ),
    "other_income": (
        "Other non-operating income",
        "income",
        ["OtherNonOperatingIncomeExpenses"],
    ),
    "pbt": ("Profit before tax", "income", ["PretaxIncome"]),
    "tax": ("Income tax expense", "income", ["TaxProvision"]),
    "net_profit": (
        "Net profit (reported basis)",
        "income",
        ["NetIncomeIncludingNoncontrollingInterests", "NetIncome"],
    ),
    "parent_profit": ("Net profit attributable to parent", "income", ["NetIncome"]),
    "cash": ("Cash & cash equivalents", "balance", ["CashAndCashEquivalents"]),
    "receivables": ("Trade receivables", "balance", ["AccountsReceivable"]),
    "inventory": ("Inventories", "balance", ["Inventory"]),
    "ppe": ("Net property, plant & equipment", "balance", ["NetPPE"]),
    "assets": ("Total assets", "balance", ["TotalAssets"]),
    "payables": ("Trade payables", "balance", ["AccountsPayable"]),
    "debt": ("Borrowings", "balance", ["TotalDebt"]),
    "liabilities": (
        "Total liabilities",
        "balance",
        ["TotalLiabilitiesNetMinorityInterest"],
    ),
    "equity": (
        "Total equity (reported basis)",
        "balance",
        ["TotalEquityGrossMinorityInterest", "StockholdersEquity"],
    ),
    "parent_equity": ("Parent equity", "balance", ["StockholdersEquity"]),
    "minorities": ("Minority equity", "balance", ["MinorityInterest"]),
    "retained_earnings": ("Retained earnings", "balance", ["RetainedEarnings"]),
    "cfo": ("Cash from operating activities", "cashflow", ["OperatingCashFlow"]),
    "cfi": ("Cash from investing activities", "cashflow", ["InvestingCashFlow"]),
    "cff": ("Cash from financing activities", "cashflow", ["FinancingCashFlow"]),
    "capex": (
        "Cash purchases of PPE",
        "cashflow",
        ["PurchaseOfPPE", "CapitalExpenditure"],
    ),
    "dividends": (
        "Cash dividends paid",
        "cashflow",
        ["CashDividendsPaid", "CommonStockDividendPaid"],
    ),
    "fx": ("FX effect on cash", "cashflow", ["EffectOfExchangeRateChanges"]),
    "cash_change": ("Reported change in cash", "cashflow", ["ChangesInCash"]),
    "cash_begin": ("Cash flow opening cash", "cashflow", ["BeginningCashPosition"]),
    "cash_end": ("Cash flow closing cash", "cashflow", ["EndCashPosition"]),
    "net_borrowing": ("Net debt cash flow", "cashflow", ["NetIssuancePaymentsOfDebt"]),
    "debt_issued": ("Cash debt issued", "cashflow", ["IssuanceOfDebt"]),
    "debt_repaid": ("Cash debt repaid", "cashflow", ["RepaymentOfDebt"]),
    "net_equity_cash_flow": (
        "Net common shares issued / (repurchased)",
        "cashflow",
        ["NetCommonStockIssuance"],
    ),
    "shares": (
        "Ordinary shares outstanding (millions)",
        "balance",
        ["OrdinarySharesNumber"],
    ),
    "share_capital": ("Share capital", "balance", ["CapitalStock", "CommonStock"]),
    "reserves": ("Reserves", "balance", []),
    "other_assets": ("Other assets (source aggregate)", "balance", []),
    "other_liabilities": ("Other liabilities (source aggregate)", "balance", []),
    "cwip": ("Capital work in progress", "balance", []),
    "investments": ("Investments", "balance", ["InvestmentsAndAdvances"]),
}


def finite(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, str):
        value = value.strip().replace(",", "").replace("\u2212", "-")
        if value in ("", "-", "—", "NA", "N/A"):
            return None
        if value.startswith("(") and value.endswith(")"):
            value = "-" + value[1:-1]
    try:
        n = float(value)
    except (ValueError, TypeError):
        return None
    return n if math.isfinite(n) else None


def ticker_name(value):
    value = value.strip().upper()
    if not re.fullmatch(r"[A-Z0-9][A-Z0-9.\-^=]{0,31}", value):
        raise ValueError("Enter a ticker such as AAPL or TCS.NS.")
    return value


def encoded_json(value):
    return json.dumps(
        value, ensure_ascii=False, sort_keys=True, indent=2, allow_nan=False
    ).encode("utf-8")


def save_snapshot(raw, cache):
    cache = Path(cache)
    cache.mkdir(parents=True, exist_ok=True)
    body = encoded_json(raw)
    digest = hashlib.sha256(body).hexdigest()
    file = cache / f"{ticker_name(raw['ticker'])}_{digest[:16]}.json"
    if not file.exists():
        file.write_bytes(body)
    return file


def latest_snapshot(ticker, cache, provider=None):
    found = []
    for path in Path(cache).glob(f"{ticker}_*.json"):
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
            if raw.get("ticker") == ticker and (
                not provider or raw.get("provider") == provider
            ):
                found.append((raw.get("retrieved_at_utc", ""), path))
        except (OSError, ValueError):
            continue
    return max(found, key=lambda x: x[0])[1] if found else None


def discover_export(ticker, folder):
    for stem in (ticker, ticker.removesuffix(".NS").removesuffix(".BO")):
        for extension in (".xlsx", ".csv"):
            path = Path(folder) / (stem + extension)
            if path.exists():
                return path
    return None


def acquire(
    ticker, cache=None, imports=None, screener=None, snapshot=None, refresh=False
):
    ticker = ticker_name(ticker)
    cache = Path(cache or BASE / "cache/snapshots")
    india = ticker.endswith((".NS", ".BO"))
    export = (
        Path(screener)
        if screener
        else discover_export(ticker, imports or BASE / "imports")
        if india
        else None
    )
    if snapshot:
        path = Path(snapshot).resolve()
        raw = json.loads(path.read_text(encoding="utf-8"))
        if raw.get("ticker") != ticker:
            raise ValueError(
                "The snapshot ticker does not match the requested company."
            )
        return raw, path, "Offline snapshot"
    if export:
        if not export.exists():
            raise ValueError(f"Screener export not found: {export}")
        original = export.read_bytes()
        digest = hashlib.sha256(original).hexdigest()
        for path in cache.glob(f"{ticker}_*.json"):
            raw = json.loads(path.read_text(encoding="utf-8"))
            if raw.get("input_sha256") == digest:
                return raw, path, "Cached Screener import"
        raw = read_screener(export, ticker)
        return raw, save_snapshot(raw, cache), "Screener company export"
    if not refresh:
        previous = latest_snapshot(ticker, cache)
        if previous:
            return (
                json.loads(previous.read_text(encoding="utf-8")),
                previous,
                "Offline cache",
            )
        # Reuse existing project fetches rather than downloading the same company again.
        legacy = [
            *(ROOT / "01_Output/History/Legacy/Companies").glob(
                f"{ticker}_*/raw_source.json"
            ),
            *(ROOT / "01_Output/Companies").glob(f"{ticker}_*/raw_source.json"),
        ]
        for path in sorted(legacy, reverse=True):
            raw = json.loads(path.read_text(encoding="utf-8"))
            if raw.get("ticker") == ticker:
                return raw, save_snapshot(raw, cache), "Reused existing raw snapshot"
    from fetch_data import download

    raw = download(ticker)
    return raw, save_snapshot(raw, cache), "New yfinance fetch"


def period(value):
    if isinstance(value, (datetime, date)):
        return (
            value.date().isoformat()
            if isinstance(value, datetime)
            else value.isoformat()
        )
    if not isinstance(value, str):
        return None
    value = value.strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}T", value):
        value = value[:10]
    if value.upper() in ("TTM", "LATEST", "CURRENT"):
        return None
    for fmt in (
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%b %Y",
        "%b-%Y",
        "%B %Y",
        "%b%Y",
        "%b-%y",
        "%b %y",
        "%Y/%m/%d",
    ):
        try:
            parsed = datetime.strptime(value, fmt)
            if fmt.startswith(("%b", "%B")):
                import calendar

                return date(
                    parsed.year,
                    parsed.month,
                    calendar.monthrange(parsed.year, parsed.month)[1],
                ).isoformat()
            return parsed.date().isoformat()
        except ValueError:
            continue
    return None


def clean_label(value):
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def read_screener(path, ticker):
    content = Path(path).read_bytes()
    sheets = {}
    if Path(path).suffix.lower() == ".xlsx":
        from openpyxl import load_workbook

        wb = load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        for sheet in wb:
            sheets[sheet.title] = [
                [x.isoformat() if isinstance(x, (date, datetime)) else x for x in row]
                for row in sheet.iter_rows(values_only=True)
            ]
        wb.close()
    elif Path(path).suffix.lower() == ".csv":
        sheets["Data Sheet"] = list(
            csv.reader(io.StringIO(content.decode("utf-8-sig")))
        )
    else:
        raise ValueError("Use a Screener company .xlsx or time-series .csv export.")
    company = ticker
    for name, rows in sheets.items():
        if clean_label(name) == "datasheet":
            for row in rows:
                if (
                    row
                    and clean_label(row[0]) == "companyname"
                    and len(row) > 1
                    and isinstance(row[1], str)
                ):
                    company = row[1].strip()
    return {
        "schema_version": 2,
        "provider": "Screener company export",
        "ticker": ticker,
        "company_name": company,
        "retrieved_at_utc": datetime.now(timezone.utc).isoformat(),
        "financial_currency": "INR",
        "basis": "Reporting basis not independently verified",
        "source_url": f"https://www.screener.in/company/{ticker.split('.')[0]}/",
        "input_name": Path(path).name,
        "input_sha256": hashlib.sha256(content).hexdigest(),
        "original_file_base64": base64.b64encode(content).decode(),
        "raw_sheets": sheets,
    }


SCREENER_MAP = {
    "income": {
        "sales": "revenue",
        "revenue": "revenue",
        "revenuefromoperations": "revenue",
        "costofgoodssold": "cogs",
        "costofsales": "cogs",
        "grossprofit": "gross_profit",
        "operatingprofit": "ebitda",
        "ebitda": "ebitda",
        "depreciation": "da",
        "depreciationandamortization": "da",
        "interest": "interest",
        "financecosts": "interest",
        "otherincome": "other_income",
        "profitbeforetax": "pbt",
        "tax": "tax",
        "taxexpense": "tax",
        "netprofit": "net_profit",
        "profitaftertax": "net_profit",
    },
    "balance": {
        "equitycapital": "share_capital",
        "sharecapital": "share_capital",
        "reserves": "reserves",
        "borrowings": "debt",
        "totalborrowings": "debt",
        "otherliabilities": "other_liabilities",
        "totalliabilities": "screener_total_funding",
        "fixedassets": "ppe",
        "netblock": "ppe",
        "cwip": "cwip",
        "investments": "investments",
        "otherassets": "other_assets",
        "totalassets": "assets",
        "tradereceivables": "receivables",
        "debtors": "receivables",
        "inventories": "inventory",
        "inventory": "inventory",
        "tradepayables": "payables",
        "creditors": "payables",
        "cashandcashequivalents": "cash",
        "cashbank": "cash",
        "cashandbank": "cash",
        "cashequivalents": "cash",
    },
    "cashflow": {
        "cashfromoperatingactivity": "cfo",
        "cashfromoperatingactivities": "cfo",
        "cashfrominvestingactivity": "cfi",
        "cashfrominvestingactivities": "cfi",
        "cashfromfinancingactivity": "cff",
        "cashfromfinancingactivities": "cff",
        "netcashflow": "cash_change",
        "capitalexpenditure": "capex",
        "purchaseoffixedassets": "capex",
        "dividendpaid": "dividends",
    },
    "ratios": {
        "debtordays": "source_dso",
        "inventorydays": "source_dio",
        "dayspayable": "source_dpo",
        "creditordays": "source_dpo",
    },
}


def normalize(raw, years=10, currency=None):
    if raw.get("provider") == "Screener company export":
        return normalize_screener(raw, years)
    if raw.get("provider") != "Yahoo Finance via yfinance":
        raise ValueError("Unsupported source snapshot.")
    if "financial" in (raw.get("sector") or "").lower():
        raise ValueError("Financial institutions require a specialist model.")
    currency = (currency or raw.get("financial_currency") or "").upper()
    if not re.fullmatch("[A-Z]{3}", currency):
        raise ValueError(
            "The source did not supply its reporting currency. Enable Fetch new data and "
            "try again; the model needs the reporting currency to use the statement "
            "values."
        )
    factor = 1e7 if currency == "INR" else 1e6
    rows = []
    lineage = []
    notes = [raw.get("basis") or "Reporting basis unverified"]
    tables = raw.get("statements", {})
    periods = sorted(
        p
        for p, v in tables.get("income", {}).items()
        if finite(v.get("TotalRevenue")) is not None
    )[-years:]
    if not periods:
        raise ValueError("No usable annual revenue observations in this snapshot.")
    for p in periods:
        date.fromisoformat(p)
        values = {}
        for key, (_label, statement, tags) in FIELDS.items():
            source = ""
            value = None
            note = ""
            choices = [(statement, tag) for tag in tags]
            if key == "da":
                choices.append(("cashflow", "DepreciationAndAmortization"))
            for section, tag in choices:
                n = finite(tables.get(section, {}).get(p, {}).get(tag))
                if n is not None:
                    value = n
                    source = f"{section}.{tag}"
                    break
            if key in ("capex", "dividends", "debt_repaid") and value is not None:
                value = abs(value)
            if (
                key in ("revenue", "cogs", "da", "interest")
                and value is not None
                and value < 0
            ):
                note = "Unexpected negative expense/revenue; review original source."
                value = None
            if key == "da" and source.startswith("cashflow"):
                note = "Cash-flow D&A may include amortization outside PPE. Review before forecasting."
            if key == "net_profit" and source.endswith(".NetIncome"):
                note = "Parent profit fallback: verify minority-interest basis against equity."
            if key == "capex" and source.endswith(".CapitalExpenditure"):
                note = "Provider aggregate CapEx may include intangible investment. Review asset scope."
            scaled = (
                value / (1e6 if key == "shares" else factor)
                if value is not None
                else None
            )
            values[key] = scaled
            lineage.append(
                {
                    "period": p,
                    "key": key,
                    "field": source,
                    "value": scaled,
                    "note": note,
                    "status": "MISSING"
                    if value is None
                    else "REVIEW"
                    if note
                    else "FETCHED",
                }
            )
        rows.append({"period": p, "values": values})
    notes += [
        (
            "Other asset/liability remainders are explicitly derived in Excel and held "
            "constant, not treated as verified operating classifications."
        ),
        "Provider statement definitions and minority-interest treatment require review.",
    ]
    if len(rows) < years:
        notes.append(
            f"Requested {years} historical years; source supplied {len(rows)}. No missing years were invented."
        )
    return {
        "ticker": raw["ticker"],
        "company": raw.get("company_name", raw["ticker"]),
        "currency": currency,
        "factor": factor,
        "unit": "INR crores" if currency == "INR" else f"{currency} millions",
        "history": rows,
        "lineage": lineage,
        "notes": notes,
        "basis": raw.get("basis", "Unverified"),
        "provider": raw["provider"],
        "retrieved": raw.get("retrieved_at_utc"),
        "source_url": raw.get("source_url"),
        "screener": False,
    }


def normalize_screener(raw, years=10):
    by_period = {}
    lineage = []
    current = "income"
    headers = {}
    all_text = []
    sheets = raw["raw_sheets"]
    # Screener templates can contain stale illustrative financial statements.
    # When present, the provider's raw Data Sheet is the sole financial source.
    canonical = {
        name: rows for name, rows in sheets.items() if clean_label(name) == "datasheet"
    }
    for sheet_name, rows in (canonical or sheets).items():
        sheet_type = clean_label(sheet_name)
        if "quarter" in sheet_type:
            continue
        current = (
            "cashflow"
            if "cashflow" in sheet_type
            else "balance"
            if "balancesheet" in sheet_type
            else "ratios"
            if "ratio" in sheet_type
            else "income"
        )
        headers = {}
        balance_part = None
        for row_number, row in enumerate(rows, 1):
            if not row:
                continue
            all_text.extend(str(x) for x in row[:3] if x is not None)
            label = next(
                (
                    str(x).strip()
                    for x in row[:3]
                    if isinstance(x, str) and x.strip() and not period(x)
                ),
                "",
            )
            name = clean_label(label)
            if name in (
                "profitloss",
                "profitandloss",
                "incomestatement",
                "annualprofitloss",
            ):
                current = "income"
                headers = {}
            elif name in ("balancesheet", "balance"):
                current = "balance"
                headers = {}
                balance_part = None
            elif name in ("cashflow", "cashflows", "cashflowstatement"):
                current = "cashflow"
                headers = {}
            elif name in ("ratios", "financialratios"):
                current = "ratios"
                headers = {}
            elif "quarter" in name:
                current = "quarterly"
                headers = {}
            dates = {i: period(x) for i, x in enumerate(row) if period(x)}
            if len(dates) >= 2:
                headers = dates
                continue
            if not headers:
                continue
            key = SCREENER_MAP.get(current, {}).get(name)
            if current == "balance":
                if name in (
                    "equitysharecapital",
                    "equitycapital",
                    "sharecapital",
                    "borrowings",
                ):
                    balance_part = "funding"
                elif name in ("netblock", "fixedassets", "fixedassetsnetblock"):
                    balance_part = "assets"
                if name == "total":
                    key = {"funding": "screener_total_funding", "assets": "assets"}.get(
                        balance_part
                    )
                key = key or {
                    "equitysharecapital": "share_capital",
                    "fixedassetsnetblock": "ppe",
                    "capitalworkinprogress": "cwip",
                    "receivables": "receivables",
                    "noofequityshares": "shares",
                }.get(name)
            if not key:
                continue
            for index, p in headers.items():
                if index >= len(row):
                    continue
                value = finite(row[index])
                target = by_period.setdefault(p, {})
                if key == "shares" and value is not None:
                    value /= 1e6
                if target.get(key) is not None and value is None:
                    continue
                if (
                    target.get(key) is not None
                    and target[key] != value
                    and value is not None
                ):
                    raise ValueError(
                        f"Conflicting Screener rows for {key} {p}; review the export layout."
                    )
                target[key] = value
                unit_note = (
                    "Source share count converted to millions."
                    if key == "shares"
                    else "Screener company-export INR-crore convention; confirm source units."
                )
                lineage.append(
                    {
                        "period": p,
                        "key": key,
                        "field": f"{sheet_name}!row {row_number}: {label}",
                        "value": value,
                        "note": unit_note,
                        "status": "MISSING" if value is None else "IMPORTED",
                    }
                )
    periods = sorted(
        p for p, values in by_period.items() if values.get("revenue") is not None
    )[-years:]
    if not periods:
        raise ValueError(
            "No annual company time series recognized. Use the company's Export to Excel "
            "file, not a screen-results list; annual headers must contain full dates or "
            "month/year."
        )
    # Exclude quarterly periods even if a custom template mixes sections.
    if len({p[:4] for p in periods}) != len(periods):
        raise ValueError(
            "Multiple annual columns share a fiscal year. Separate quarterly/TTM data from annual statements."
        )
    text = " ".join(all_text).lower()
    if any(
        re.search(
            r"(?:figures|amounts|units|currency|inr|rs\.?)[^\n]{0,30}\b(?:lakhs?|lacs?|millions?|billions?|usd|eur)\b",
            cell.lower(),
        )
        for cell in all_text
    ):
        raise ValueError(
            "This Screener layout declares units other than INR crores. Convert the "
            "annual monetary figures to INR crores before importing; share counts remain "
            "absolute shares."
        )
    basis = (
        "Consolidated"
        if "consolidated" in text and "standalone" not in text
        else "Standalone"
        if "standalone" in text and "consolidated" not in text
        else raw.get("basis", "Unverified")
    )
    notes = [
        (
            "Screener imported values use the company's export convention of INR crores; "
            "verify units and reporting basis."
        ),
        (
            "Screener Total Liabilities can include equity: retained separately as total "
            "funding, never substituted for liabilities excluding equity."
        ),
        (
            "Operating Profit in Screener is treated as EBITDA. Expenses are not assumed "
            "to equal COGS; working-capital days remain unavailable without an "
            "appropriate cost basis."
        ),
        (
            "Incomplete detailed working capital and asset classifications remain "
            "missing. Importing about ten years does not guarantee every model field is "
            "disclosed."
        ),
    ]
    if canonical:
        notes.append(
            "Only the raw Data Sheet was imported. Presentation/template tabs and "
            "supplemental Cash Flow Data were excluded to avoid mixing stale examples or "
            "reporting periods."
        )
    notes.append(
        "Where Screener EBITDA/EBIT is absent, Excel reconciles EBIT as PBT + "
        "interest - other income and EBITDA as EBIT + D&A. Review exceptional items "
        "before using these as operating margins."
    )
    if len(periods) < years:
        notes.append(
            f"Only {len(periods)} usable annual periods were available; no additional history was fabricated."
        )
    return {
        "ticker": raw["ticker"],
        "company": raw.get("company_name", raw["ticker"]),
        "currency": "INR",
        "factor": 1e7,
        "unit": "INR crores",
        "history": [{"period": p, "values": by_period[p]} for p in periods],
        "lineage": [x for x in lineage if x["period"] in periods],
        "notes": notes,
        "basis": basis,
        "provider": raw["provider"],
        "retrieved": raw.get("retrieved_at_utc"),
        "source_url": raw.get("source_url"),
        "screener": True,
    }
