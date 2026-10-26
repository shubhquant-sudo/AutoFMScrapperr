import math
import re
from datetime import datetime, timezone
from pathlib import Path

BASE = Path(__file__).resolve().parent


class FetchError(ValueError):
    pass


def symbol(value):
    value = value.strip().upper()
    if not re.fullmatch(r"[A-Z0-9][A-Z0-9.\^=\-]{0,31}", value):
        raise FetchError("Enter a Yahoo ticker such as RELIANCE.NS, TCS.NS or AAPL")
    return value


def number(value):
    if isinstance(value, bool) or value is None:
        return None
    try:
        result = float(value)
    except (TypeError, ValueError):
        return None
    return result if math.isfinite(result) else None


def download(ticker):
    try:
        import yfinance as yf
    except ImportError as error:
        raise FetchError(
            "Install dependencies first: python -m pip install -r requirements.txt"
        ) from error
    cache = BASE / ".cache/yfinance"
    cache.mkdir(parents=True, exist_ok=True)
    yf.set_tz_cache_location(str(cache))
    company = yf.Ticker(ticker)
    tables = {}
    for name, method in [
        ("income", company.get_income_stmt),
        ("balance", company.get_balance_sheet),
        ("cashflow", company.get_cash_flow),
    ]:
        try:
            frame = method(pretty=False, freq="yearly")
        except Exception as error:
            raise FetchError(f"Yahoo {name} download failed: {error}") from error
        periods = {}
        if frame is not None and not frame.empty:
            if frame.columns.duplicated().any() or frame.index.duplicated().any():
                raise FetchError(f"Yahoo returned duplicate {name} dates or line items")
            for column in frame.columns:
                period = column.date().isoformat()
                periods[period] = {
                    str(key): number(value) for key, value in frame[column].items()
                }
        tables[name] = periods
    if not tables["income"]:
        raise FetchError(
            "Yahoo returned no annual income statement. Check the ticker, internet "
            "connection or rate limit and retry later."
        )
    try:
        info = company.get_info() or {}
    except Exception:
        info = {}  # Missing financial currency still blocks unless explicitly supplied.
    return {
        "schema_version": 1,
        "provider": "Yahoo Finance via yfinance",
        "ticker": ticker,
        "retrieved_at_utc": datetime.now(timezone.utc).isoformat(),
        "company_name": info.get("longName") or info.get("shortName") or ticker,
        "financial_currency": info.get("financialCurrency"),
        "sector": info.get("sector"),
        "industry": info.get("industry"),
        "basis": "Provider annual statements; consolidated/standalone basis not independently verified",
        "source_url": f"https://finance.yahoo.com/quote/{ticker}/financials/",
        "statements": tables,
    }
