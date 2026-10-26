"""A small linked three-statement model. All business calculations live in Excel."""

from functools import partial

from openpyxl.utils import get_column_letter as column_letter

from .data import FIELDS, clean_label

NUM = '#,##0.0;(#,##0.0);"-"'
PCT = '0.0%;(0.0%);"-"'
PRICE = '#,##0.00;(#,##0.00);"-"'
METHOD = "https://www.cfainstitute.org/insights/professional-learning/refresher-readings/2026/free-cash-flow-valuation"
DRIVERS = [
    (
        "growth",
        "Revenue growth",
        PCT,
        "Enter the annual change in sales. For 5%, type 5%.",
        {"min": -1, "strict": True},
    ),
    (
        "gross_margin",
        "Gross profit margin",
        PCT,
        "Gross profit divided by sales. This sets cost of sales.",
        {"max": 1},
    ),
    (
        "ebitda_margin",
        "EBITDA margin",
        PCT,
        "Profit before depreciation, interest and tax, divided by sales. Must not exceed gross margin.",
        {},
    ),
    (
        "tax_rate",
        "Tax rate on positive profit",
        PCT,
        "Tax as a share of positive profit. Losses receive no automatic tax credit.",
        {"min": 0, "max": 1},
    ),
    (
        "receivable_days",
        "Customer payment days",
        "0.0",
        "Average days before customers pay. Receivables = sales / 365 × days.",
        {"min": 0},
    ),
    (
        "inventory_days",
        "Inventory holding days",
        "0.0",
        "Inventory = cost of sales / 365 × days. Enter 0 if the business holds none.",
        {"min": 0},
    ),
    (
        "payable_days",
        "Supplier payment days",
        "0.0",
        "Payables = cost of sales / 365 × days. Enter 0 if not used.",
        {"min": 0},
    ),
    (
        "capex_rate",
        "Capital spending / revenue",
        PCT,
        "Cash spent on fixed assets, divided by sales. Enter 0 for no spending.",
        {"min": 0},
    ),
    (
        "depreciation_rate",
        "Depreciation / opening fixed assets",
        PCT,
        "Rate applied to opening net fixed assets. New spending starts depreciating next year.",
        {"min": 0, "max": 1},
    ),
    (
        "interest_rate",
        "Interest rate on opening debt",
        PCT,
        "Annual financing rate. Interest uses opening debt, avoiding a circular formula.",
        {"min": 0},
    ),
    (
        "repayment_rate",
        "Debt repayment / opening debt",
        PCT,
        "Share of opening debt repaid during the year. Enter 0 for no repayment.",
        {"min": 0, "max": 1},
    ),
    (
        "new_borrowing",
        "New borrowing",
        NUM,
        "Amount of new debt, in the workbook currency units. Enter 0 if none.",
        {"min": 0},
    ),
    (
        "payout_rate",
        "Dividend payout / positive profit",
        PCT,
        "Share of positive net profit paid as dividends. Enter 0 to retain all profit.",
        {"min": 0, "max": 1},
    ),
    (
        "other_income",
        "Other income (excluding debt interest)",
        NUM,
        (
            "Net non-operating income or costs. Exclude the debt interest already modeled"
            " above. Enter 0 if none."
        ),
        {},
    ),
    (
        "equity_cash_flow",
        "Shares issued / (repurchased)",
        NUM,
        (
            "Cash from new shares is positive; buybacks are negative. Enter 0 if none. "
            "Review the DCF share count separately."
        ),
        {},
    ),
]
VALUATION = [
    (
        "wacc",
        "Discount rate (WACC)",
        PCT,
        "Required return for debt and equity together. Must be positive.",
        {"min": 0, "strict": True},
    ),
    (
        "terminal_growth",
        "Long-run cash flow growth",
        PCT,
        "Growth after year five. Must be greater than -100% and below WACC.",
        {"min": -1, "strict": True},
    ),
    (
        "shares",
        "Shares outstanding (millions)",
        NUM,
        "Enter your current fully diluted share count in millions. History shows the reported count.",
        {"min": 0, "strict": True},
    ),
    (
        "claims",
        "Other claims to deduct",
        NUM,
        (
            "Claims on value other than the debt already shown. Include minority "
            "interests if relevant; enter 0 if none."
        ),
        {"min": 0},
    ),
    (
        "nonoperating",
        "Other non-operating assets to add",
        NUM,
        "Eligible assets not counted in cash or operating value. Do not add goodwill; enter 0 if none.",
        {"min": 0},
    ),
    (
        "price",
        "Current share price (optional)",
        PRICE,
        "Only needed for upside/downside. Use the same currency as the per-share value.",
        {"min": 0, "strict": True},
    ),
]


def ref(sheet, row, column_number):
    return f"'{sheet}'!{column_letter(column_number)}{row}"


def guard(expression, dependencies, condition="TRUE"):
    return f'=IF(COUNT({",".join(dependencies)})={len(dependencies)},IF({condition},{expression},""),"")'


class WorksheetLayout:
    """Collect cells and formatting for the Excel writer."""

    def __init__(self, specification):
        self.specification = specification
        self.worksheets = {}

    def add_sheet(self, name, rows, notes_col=None):
        worksheet = {
            "name": name,
            "rows": rows,
            "cells": [],
            "sections": [],
            "totals": [],
            "notes": [],
            "notes_col": notes_col,
            "merges": [],
        }
        self.worksheets[name] = worksheet
        self.specification["sheets"].append(worksheet)
        return worksheet

    def write_cell(self, sheet_name, row, column, value, number_format=NUM, kind=None):
        worksheet = self.worksheets[sheet_name]
        worksheet["cells"].append(
            {
                "row": row,
                "col": column,
                "value": value,
                "format": number_format,
                "kind": kind,
            }
        )
        worksheet["rows"] = max(worksheet["rows"], row + 1)
        return ref(sheet_name, row, column)

    def write_formula(
        self,
        sheet_name,
        row,
        column,
        expression,
        dependencies,
        condition="TRUE",
        number_format=NUM,
    ):
        return self.write_cell(
            sheet_name,
            row,
            column,
            guard(expression, dependencies, condition),
            number_format,
        )

    def write_label(self, sheet_name, row, text):
        self.write_cell(sheet_name, row, 2, text, "General", "label")

    def write_note(self, sheet_name, row, text):
        worksheet = self.worksheets[sheet_name]
        worksheet["notes"].append([row, text])
        worksheet["rows"] = max(worksheet["rows"], row + 2)

    def add_section(self, sheet_name, row, text):
        self.worksheets[sheet_name]["sections"].append([row, text])

    def mark_totals(self, sheet_name, *rows):
        self.worksheets[sheet_name]["totals"] += rows


def _write_history(layout, data, raw):
    """Keep editable history separate from the original download."""
    history = data["history"]
    model_spec = layout.specification
    keys = list(FIELDS) + sorted(
        {k for row in history for k in row["values"]} - set(FIELDS)
    )
    history_rows = {k: 9 + index for index, k in enumerate(keys)}
    by_period = {row["period"]: 4 + index for index, row in enumerate(history)}
    layout.write_note(
        "History",
        4,
        (
            "All available annual history is here. Green numbers came from the saved "
            "source. Yellow blanks are missing. Correct a figure here only after checking"
            " the original statements. Raw statements and downloaded files remain "
            "unchanged."
        ),
    )
    layout.write_cell("History", 7, 2, "Reported item", "General", "header")
    layout.write_cell("History", 7, 3, "Source field", "General", "header")
    for period, column_number in by_period.items():
        layout.write_cell("History", 7, column_number, period, "General", "header")
    for key, row_number in history_rows.items():
        layout.write_label("History", row_number, FIELDS.get(key, (key,))[0])
        layout.write_cell("History", row_number, 3, key, "General", "source_label")
        for record in history:
            column_number = by_period[record["period"]]
            value = record["values"].get(key)
            layout.write_cell(
                "History",
                row_number,
                column_number,
                value,
                NUM,
                "source" if value is not None else "missing_source",
            )
            model_spec["history_inputs"][record["period"] + "." + key] = {
                "sheet": "History",
                "cell": f"{column_letter(column_number)}{row_number}",
                "missing": value is None,
            }
    layout.write_note(
        "History",
        len(keys) + 12,
        "Source: "
        + str(data.get("source_url") or raw.get("source_url") or raw["provider"])
        + ". Retrieved "
        + str(data.get("retrieved") or raw.get("retrieved_at_utc", "not recorded"))
        + ". Reporting basis: "
        + str(data.get("basis", "unverified"))
        + ".",
    )
    return history_rows, by_period


def _write_raw_statements(layout, raw):
    """Retain original source lines without using them as forecast assumptions."""
    row_number = 7
    if raw.get("statements"):
        order = [
            k for k in ("income", "balance", "cashflow") if k in raw["statements"]
        ] + sorted(set(raw["statements"]) - {"income", "balance", "cashflow"})
        for kind in order:
            periods = raw["statements"][kind]
            dates = sorted(periods)
            layout.add_section(
                "Raw statements",
                row_number,
                {
                    "income": "Reported income statement",
                    "balance": "Reported balance sheet",
                    "cashflow": "Reported cash flow",
                }.get(kind, kind),
            )
            row_number += 2
            layout.write_cell(
                "Raw statements",
                row_number,
                2,
                "Original provider field",
                "General",
                "header",
            )
            for column_number, date in enumerate(dates, 4):
                layout.write_cell(
                    "Raw statements",
                    row_number,
                    column_number,
                    date,
                    "General",
                    "header",
                )
            fields = sorted({k for values in periods.values() for k in values})
            for key in fields:
                row_number += 1
                layout.write_cell(
                    "Raw statements", row_number, 2, key, "General", "source_label"
                )
                for column_number, date in enumerate(dates, 4):
                    layout.write_cell(
                        "Raw statements",
                        row_number,
                        column_number,
                        periods[date].get(key),
                        "General",
                        "raw_source",
                    )
            row_number += 4
    else:
        raw_sheets = raw.get("raw_sheets", {})
        canonical = {
            name: rows
            for name, rows in raw_sheets.items()
            if clean_label(name) == "datasheet"
        }
        for name, rows in (canonical or raw_sheets).items():
            layout.add_section("Raw statements", row_number, name)
            row_number += 2
            for row in rows:
                for column_number, value in enumerate(row, 2):
                    if value is not None:
                        layout.write_cell(
                            "Raw statements",
                            row_number,
                            column_number,
                            value,
                            "General",
                            "raw_source",
                        )
                row_number += 1
            row_number += 3
    layout.write_note(
        "Raw statements",
        4,
        (
            "Original provider values, before unit conversion or mapping. These may use "
            "currency units rather than the millions/crores in the model. The downloaded "
            "JSON or imported export is also saved with this workbook."
        ),
    )


def _write_assumptions(layout, data, last_actual_column):
    """Start with blank forecast inputs and last-actual reference ratios."""
    history = data["history"]
    model_spec = layout.specification
    layout.write_note(
        "Assumptions",
        4,
        (
            "Start with History and check the latest actual figures. Then fill the yellow"
            " forecast cells below. Enter 0 when you intend none. A blank means missing, "
            "so the affected forecast stays blank."
        ),
    )
    layout.write_cell("Assumptions", 7, 2, "Forecast driver", "General", "header")
    layout.write_cell("Assumptions", 7, 4, "Last actual", "General", "header")
    layout.write_cell("Assumptions", 7, 12, "What to enter", "General", "header")
    for index, column_number in enumerate(range(6, 11)):
        layout.write_cell(
            "Assumptions",
            7,
            column_number,
            int(history[-1]["period"][:4]) + index + 1,
            '0" forecast"',
            "header",
        )
    assumption_rows = {}
    for row_number, (key, title, number_format, why, rule) in enumerate(DRIVERS, 9):
        layout.write_label("Assumptions", row_number, title)
        layout.write_cell("Assumptions", row_number, 12, why, "General", "explanation")
        assumption_rows[key] = row_number
        for column_number in range(6, 11):
            layout.write_cell(
                "Assumptions", row_number, column_number, None, number_format, "input"
            )
            model_spec["inputs"][key + "." + str(column_number - 5)] = {
                "sheet": "Assumptions",
                "cell": f"{column_letter(column_number)}{row_number}",
                "rule": rule,
                "optional": False,
            }
        benchmark = ref("Statements", 120 + row_number - 9, last_actual_column)
        layout.write_formula(
            "Assumptions",
            row_number,
            4,
            benchmark,
            [benchmark],
            number_format=number_format,
        )
    layout.write_label("Assumptions", 25, "Unfilled forecast inputs")
    for column_number in range(6, 11):
        layout.write_cell(
            "Assumptions",
            25,
            column_number,
            f"=15-COUNT({column_letter(column_number)}9:{column_letter(column_number)}23)",
            "0",
            "count",
        )
    layout.add_section(
        "Assumptions", 29, "DCF inputs. Enter these once, after the forecast."
    )
    for row_number, (key, title, number_format, why, rule) in enumerate(VALUATION, 31):
        if key == "price":
            why = (
                "Optional current price in "
                + data["currency"]
                + " per share. Leave blank if you do not need upside/downside."
            )
        layout.write_label("Assumptions", row_number, title)
        layout.write_cell("Assumptions", row_number, 6, None, number_format, "input")
        layout.write_cell("Assumptions", row_number, 12, why, "General", "explanation")
        model_spec["inputs"][key] = {
            "sheet": "Assumptions",
            "cell": f"F{row_number}",
            "rule": rule,
            "optional": key == "price",
        }
    layout.write_note(
        "Assumptions",
        39,
        (
            "To keep assumptions the same each year, fill F9:F23, select F9:J23 and press"
            " Ctrl+R. Then adjust individual years if needed. The last-actual column is a"
            " reference, not a suggested forecast."
        ),
    )
    layout.write_note(
        "Assumptions",
        43,
        (
            "How the basic forecast works: other assets and other liabilities stay at "
            "their latest actual amounts. Debt and fixed assets roll forward. There are "
            "no acquisitions, disposals, stock compensation or new tax-loss schedules. "
            "Negative cash remains visible; borrowing is never added automatically to "
            "make the model balance."
        ),
    )
    return assumption_rows


def make_model(data, raw):
    model_spec = {
        "company": data["company"],
        "ticker": data["ticker"],
        "unit": data["unit"],
        "currency": data["currency"],
        "basis": data.get("basis", "Unverified"),
        "sheets": [],
        "inputs": {},
        "history_inputs": {},
        "outputs": {},
        "checks": [],
        "periods": [row_number["period"] for row_number in data["history"]],
    }
    layout = WorksheetLayout(model_spec)

    layout.add_sheet("Assumptions", 47, 12)
    layout.add_sheet("Statements", 138, 14)
    layout.add_sheet("DCF", 55, 10)
    layout.add_sheet("History", 65, 3)
    layout.add_sheet("Raw statements", 10)
    history = data["history"]
    visible = history[-3:]
    historical_columns = list(range(4, 4 + len(visible)))
    last_actual_column = historical_columns[-1]
    forecast_columns = list(range(last_actual_column + 2, last_actual_column + 7))
    model_spec.update(
        historical_columns=historical_columns,
        forecast_columns=forecast_columns,
        last_actual_column=last_actual_column,
    )
    # Original observations, including missing values, remain separate from the model.
    history_rows, by_period = _write_history(layout, data, raw)

    def history_ref(key, period):
        return ref("History", history_rows[key], by_period[period])

    # Keep every raw provider line and period, not just the normalized fields.
    _write_raw_statements(layout, raw)
    # Instructions sit beside the inputs, not on a decorative cover.
    assumption_rows = _write_assumptions(layout, data, last_actual_column)
    # Three statements, kept together so the links are easy to follow.
    labels = {
        9: "Revenue",
        10: "Cost of sales",
        11: "Gross profit",
        12: "Cash operating costs (derived)",
        13: "EBITDA",
        14: "Depreciation and amortization",
        15: "Operating profit (EBIT)",
        16: "Net finance and other costs",
        17: "Profit before tax",
        18: "Income tax",
        19: "Net profit",
        31: "Cash and cash equivalents",
        32: "Trade receivables",
        33: "Inventory",
        34: "Net fixed assets",
        35: "Other assets (derived)",
        36: "Total assets",
        38: "Trade payables",
        39: "Total debt",
        40: "Other liabilities (derived)",
        41: "Total liabilities",
        42: "Total equity",
        43: "Liabilities and equity",
        58: "Net profit",
        59: "Depreciation added back",
        60: "Other operating cash adjustments",
        61: "Increase in working capital",
        62: "Cash from operations",
        63: "Capital spending (outflow)",
        64: "Other investing cash flow",
        65: "Cash from investing",
        66: "Net borrowing / (repayment)",
        67: "Dividends paid (outflow)",
        68: "Share financing / other cash flows",
        69: "Cash from financing",
        70: "Other cash movement (derived)",
        71: "Net change in cash",
        72: "Opening cash",
        73: "Closing cash",
        82: "Opening fixed assets",
        83: "Capital spending",
        84: "Depreciation",
        85: "Closing fixed assets",
        86: "Operating working capital",
        87: "Opening debt",
        88: "Forecast inputs valid",
        89: "Opening figures ready",
        90: "Closing debt",
        107: "Balance sheet difference",
        108: "Profit reconciliation difference",
        109: "Cash flow reconciliation difference",
        110: "Cash balance review",
    }
    for row_number, text in labels.items():
        layout.write_label("Statements", row_number, text)
    for row_number, title in [
        (8, "Income statement"),
        (29, "Balance sheet"),
        (55, "Cash flow statement"),
        (80, "Supporting calculations"),
        (119, "Historical driver reference"),
        (105, "Review checks"),
    ]:
        layout.add_section("Statements", row_number, title)
    layout.write_cell("Statements", 7, 2, "Financial year", "General", "header")
    for record, column_number in zip(visible, historical_columns, strict=False):
        layout.write_cell(
            "Statements",
            7,
            column_number,
            int(record["period"][:4]),
            '0" actual"',
            "header",
        )
    for index, column_number in enumerate(forecast_columns):
        layout.write_cell(
            "Statements",
            7,
            column_number,
            int(history[-1]["period"][:4]) + index + 1,
            '0" forecast"',
            "header",
        )
    layout.write_note(
        "Statements",
        4,
        (
            "Read down each statement. Black numbers calculate within this sheet; green "
            "numbers link to another sheet. Forecasts appear one year at a time after the"
            " inputs and opening figures are complete. Review checks are at the bottom."
        ),
    )
    for record, column_number in zip(visible, historical_columns, strict=False):
        period = record["period"]
        source_ref = partial(history_ref, period=period)
        statement_ref = partial(ref, "Statements", column_number=column_number)

        def write_actual(row, key, column=column_number, source=source_ref):
            layout.write_formula("Statements", row, column, source(key), [source(key)])

        for row, key in [
            (9, "revenue"),
            (10, "cogs"),
            (14, "da"),
            (17, "pbt"),
            (18, "tax"),
            (19, "net_profit"),
            (31, "cash"),
            (32, "receivables"),
            (33, "inventory"),
            (34, "ppe"),
            (36, "assets"),
            (38, "payables"),
            (39, "debt"),
            (62, "cfo"),
            (63, "capex"),
            (65, "cfi"),
            (67, "dividends"),
        ]:
            write_actual(row, key)
        fallback = (
            f'IF(COUNT({source_ref("pbt")},{source_ref("interest")},{source_ref("other_income")})=3,{source_ref("pbt")}+{source_ref("interest")}-{source_ref("other_income")},"")'
            if data.get("screener")
            else '""'
        )
        layout.write_cell(
            "Statements",
            15,
            column_number,
            f"=IF(ISNUMBER({source_ref('ebit')}),{source_ref('ebit')},IF(COUNT({source_ref('ebitda')},{source_ref('da')})=2,{source_ref('ebitda')}-{source_ref('da')},{fallback}))",
        )
        layout.write_cell(
            "Statements",
            42,
            column_number,
            f'=IF(ISNUMBER({source_ref("equity")}),{source_ref("equity")},IF(COUNT({source_ref("share_capital")},{source_ref("reserves")})=2,{source_ref("share_capital")}+{source_ref("reserves")},""))',
        )
        layout.write_cell(
            "Statements",
            41,
            column_number,
            f'=IF(ISNUMBER({source_ref("liabilities")}),{source_ref("liabilities")},IF(COUNT({source_ref("debt")},{source_ref("other_liabilities")})=2,{source_ref("debt")}+{source_ref("other_liabilities")},""))',
        )
        for row, expression, dependencies in [
            (11, f"{statement_ref(9)}-{statement_ref(10)}", [9, 10]),
            (13, f"{statement_ref(15)}+{statement_ref(14)}", [15, 14]),
            (12, f"{statement_ref(11)}-{statement_ref(13)}", [11, 13]),
            (16, f"{statement_ref(15)}-{statement_ref(17)}", [15, 17]),
            (
                35,
                f"{statement_ref(36)}-SUM({statement_ref(31)}:{statement_ref(34)})",
                [36, 31, 32, 33, 34],
            ),
            (
                40,
                f"{statement_ref(41)}-{statement_ref(38)}-{statement_ref(39)}",
                [41, 38, 39],
            ),
            (43, f"{statement_ref(41)}+{statement_ref(42)}", [41, 42]),
            (
                86,
                f"{statement_ref(32)}+{statement_ref(33)}-{statement_ref(38)}",
                [32, 33, 38],
            ),
            (58, statement_ref(19), [19]),
            (59, statement_ref(14), [14]),
            (64, f"{statement_ref(65)}+{statement_ref(63)}", [65, 63]),
            (73, statement_ref(31), [31]),
        ]:
            layout.write_formula(
                "Statements",
                row,
                column_number,
                expression,
                [statement_ref(row_number) for row_number in dependencies],
            )
        write_actual(69, "cff")
        previous = history.index(record) - 1
        if previous >= 0:
            prior = history[previous]["period"]
            previous_ref = partial(history_ref, period=prior)
            layout.write_formula(
                "Statements",
                72,
                column_number,
                previous_ref("cash"),
                [previous_ref("cash")],
            )
            dependencies = [
                source_ref(k) for k in ("receivables", "inventory", "payables")
            ] + [previous_ref(k) for k in ("receivables", "inventory", "payables")]
            layout.write_formula(
                "Statements",
                61,
                column_number,
                f"{statement_ref(86)}-({previous_ref('receivables')}+{previous_ref('inventory')}-{previous_ref('payables')})",
                dependencies,
            )
            layout.write_formula(
                "Statements",
                71,
                column_number,
                f"{statement_ref(73)}-{statement_ref(72)}",
                [statement_ref(73), statement_ref(72)],
            )
            layout.write_formula(
                "Statements",
                70,
                column_number,
                f"{statement_ref(71)}-{statement_ref(62)}-{statement_ref(65)}-{statement_ref(69)}",
                [statement_ref(row_number) for row_number in (71, 62, 65, 69)],
            )
        else:
            layout.write_formula(
                "Statements",
                72,
                column_number,
                source_ref("cash_begin"),
                [source_ref("cash_begin")],
            )
            layout.write_formula(
                "Statements",
                71,
                column_number,
                f"{statement_ref(73)}-{statement_ref(72)}",
                [statement_ref(73), statement_ref(72)],
            )
            layout.write_formula(
                "Statements",
                70,
                column_number,
                f"{statement_ref(71)}-{statement_ref(62)}-{statement_ref(65)}-{statement_ref(69)}",
                [statement_ref(row_number) for row_number in (71, 62, 65, 69)],
            )
        layout.write_formula(
            "Statements",
            60,
            column_number,
            f"{statement_ref(62)}-{statement_ref(58)}-{statement_ref(59)}+{statement_ref(61)}",
            [statement_ref(row_number) for row_number in (62, 58, 59, 61)],
        )
        # Financing detail uses cash-flow observations, not changes in book debt.
        write_actual(66, "net_borrowing")
        layout.write_formula(
            "Statements",
            68,
            column_number,
            f"{statement_ref(69)}-{statement_ref(66)}+{statement_ref(67)}",
            [statement_ref(69), statement_ref(66), statement_ref(67)],
        )
        dependencies = [
            statement_ref(9),
            statement_ref(10),
            statement_ref(11),
            statement_ref(13),
        ]
        benchmarks = {
            1: (
                f"{statement_ref(11)}/{statement_ref(9)}",
                [statement_ref(11), statement_ref(9)],
                f"{statement_ref(9)}>0",
            ),
            2: (
                f"{statement_ref(13)}/{statement_ref(9)}",
                [statement_ref(13), statement_ref(9)],
                f"{statement_ref(9)}>0",
            ),
            3: (
                f"{statement_ref(18)}/{statement_ref(17)}",
                [statement_ref(18), statement_ref(17)],
                f"{statement_ref(17)}>0",
            ),
            4: (
                f"{statement_ref(32)}/{statement_ref(9)}*365",
                [statement_ref(32), statement_ref(9)],
                f"{statement_ref(9)}>0",
            ),
            5: (
                f"{statement_ref(33)}/{statement_ref(10)}*365",
                [statement_ref(33), statement_ref(10)],
                f"{statement_ref(10)}>0",
            ),
            6: (
                f"{statement_ref(38)}/{statement_ref(10)}*365",
                [statement_ref(38), statement_ref(10)],
                f"{statement_ref(10)}>0",
            ),
            7: (
                f"{statement_ref(63)}/{statement_ref(9)}",
                [statement_ref(63), statement_ref(9)],
                f"{statement_ref(9)}>0",
            ),
            12: (
                f"{statement_ref(67)}/{statement_ref(19)}",
                [statement_ref(67), statement_ref(19)],
                f"{statement_ref(19)}>0",
            ),
        }
        if previous >= 0:
            benchmarks[0] = (
                f"{statement_ref(9)}/{previous_ref('revenue')}-1",
                [statement_ref(9), previous_ref("revenue")],
                f"{previous_ref('revenue')}>0",
            )
            benchmarks[8] = (
                f"{statement_ref(14)}/{previous_ref('ppe')}",
                [statement_ref(14), previous_ref("ppe")],
                f"{previous_ref('ppe')}>0",
            )
            benchmarks[9] = (
                f"{source_ref('interest')}/{previous_ref('debt')}",
                [source_ref("interest"), previous_ref("debt")],
                f"{previous_ref('debt')}>0",
            )
            benchmarks[10] = (
                f"{source_ref('debt_repaid')}/{previous_ref('debt')}",
                [source_ref("debt_repaid"), previous_ref("debt")],
                f"{previous_ref('debt')}>0",
            )
        for index, (expression, dependencies, condition) in benchmarks.items():
            layout.write_formula(
                "Statements",
                120 + index,
                column_number,
                expression,
                dependencies,
                condition,
                DRIVERS[index][2],
            )
        layout.write_formula(
            "Statements",
            133,
            column_number,
            f"{statement_ref(17)}-{statement_ref(15)}+{source_ref('interest')}",
            [statement_ref(17), statement_ref(15), source_ref("interest")],
        )
        layout.write_formula(
            "Statements",
            131,
            column_number,
            source_ref("debt_issued"),
            [source_ref("debt_issued")],
        )
        layout.write_formula(
            "Statements",
            134,
            column_number,
            source_ref("net_equity_cash_flow"),
            [source_ref("net_equity_cash_flow")],
        )
    # Forecasts use opening balances for debt interest and depreciation.
    for index, column_number in enumerate(forecast_columns):
        assumption_column = 6 + index

        def assumption_ref(key, column=assumption_column):
            return ref("Assumptions", assumption_rows[key], column)

        statement_ref = partial(ref, "Statements", column_number=column_number)
        previous_column = last_actual_column if index == 0 else column_number - 1
        previous_ref = partial(ref, "Statements", column_number=previous_column)
        inputs = [assumption_ref(key) for key, *_ in DRIVERS]
        conditions = [
            f"{assumption_ref('growth')}>-1",
            f"{assumption_ref('gross_margin')}<=1",
            f"{assumption_ref('ebitda_margin')}<={assumption_ref('gross_margin')}",
            f"{assumption_ref('tax_rate')}>=0",
            f"{assumption_ref('tax_rate')}<=1",
        ]
        for key in (
            "receivable_days",
            "inventory_days",
            "payable_days",
            "capex_rate",
            "depreciation_rate",
            "interest_rate",
            "repayment_rate",
            "new_borrowing",
            "payout_rate",
        ):
            conditions.append(assumption_ref(key) + ">=0")
        for key in ("depreciation_rate", "repayment_rate", "payout_rate"):
            conditions.append(assumption_ref(key) + "<=1")
        layout.write_cell(
            "Statements",
            88,
            column_number,
            f"=IF(COUNT({','.join(inputs)})=15,AND({','.join(conditions)}),FALSE)",
            "General",
        )
        opening = [
            previous_ref(row_number)
            for row_number in (9, 31, 34, 35, 36, 39, 40, 42, 43, 86)
        ]
        layout.write_cell(
            "Statements",
            89,
            column_number,
            f"=IF(AND({statement_ref(88)},COUNT({','.join(opening)})=10),AND({previous_ref(9)}>0,{previous_ref(34)}>=0,{previous_ref(35)}>=0,{previous_ref(39)}>=0,{previous_ref(40)}>=0,ABS({previous_ref(36)}-{previous_ref(43)})<=0.5),FALSE)",
            "General",
        )

        def write_forecast(
            row, expression, column=column_number, current_ref=statement_ref
        ):
            layout.write_cell(
                "Statements",
                row,
                column,
                f'=IF({current_ref(89)},{expression},"")',
            )

        expression = {
            9: f"{previous_ref(9)}*(1+{assumption_ref('growth')})",
            10: f"{statement_ref(9)}*(1-{assumption_ref('gross_margin')})",
            11: f"{statement_ref(9)}-{statement_ref(10)}",
            12: f"{statement_ref(11)}-{statement_ref(13)}",
            13: f"{statement_ref(9)}*{assumption_ref('ebitda_margin')}",
            14: f"{previous_ref(34)}*{assumption_ref('depreciation_rate')}",
            15: f"{statement_ref(13)}-{statement_ref(14)}",
            16: f"{previous_ref(39)}*{assumption_ref('interest_rate')}-{assumption_ref('other_income')}",
            17: f"{statement_ref(15)}-{statement_ref(16)}",
            18: f"MAX(0,{statement_ref(17)})*{assumption_ref('tax_rate')}",
            19: f"{statement_ref(17)}-{statement_ref(18)}",
            31: statement_ref(73),
            32: f"{statement_ref(9)}/365*{assumption_ref('receivable_days')}",
            33: f"{statement_ref(10)}/365*{assumption_ref('inventory_days')}",
            34: f"{previous_ref(34)}+{statement_ref(63)}-{statement_ref(14)}",
            35: previous_ref(35),
            36: f"SUM({statement_ref(31)}:{statement_ref(35)})",
            38: f"{statement_ref(10)}/365*{assumption_ref('payable_days')}",
            39: f"{previous_ref(39)}+{assumption_ref('new_borrowing')}-{previous_ref(39)}*{assumption_ref('repayment_rate')}",
            40: previous_ref(40),
            41: f"SUM({statement_ref(38)}:{statement_ref(40)})",
            42: f"{previous_ref(42)}+{statement_ref(19)}-{statement_ref(67)}+{assumption_ref('equity_cash_flow')}",
            43: f"{statement_ref(41)}+{statement_ref(42)}",
            58: statement_ref(19),
            59: statement_ref(14),
            60: "0",
            61: f"{statement_ref(86)}-{previous_ref(86)}",
            62: f"{statement_ref(58)}+{statement_ref(59)}-{statement_ref(61)}",
            63: f"{statement_ref(9)}*{assumption_ref('capex_rate')}",
            64: "0",
            65: f"-{statement_ref(63)}",
            66: f"{assumption_ref('new_borrowing')}-{previous_ref(39)}*{assumption_ref('repayment_rate')}",
            67: f"MAX(0,{statement_ref(19)})*{assumption_ref('payout_rate')}",
            68: assumption_ref("equity_cash_flow"),
            69: f"{statement_ref(66)}-{statement_ref(67)}+{statement_ref(68)}",
            70: "0",
            71: f"{statement_ref(62)}+{statement_ref(65)}+{statement_ref(69)}",
            72: previous_ref(31),
            73: f"{statement_ref(72)}+{statement_ref(71)}",
            82: previous_ref(34),
            83: statement_ref(63),
            84: statement_ref(14),
            85: statement_ref(34),
            86: f"{statement_ref(32)}+{statement_ref(33)}-{statement_ref(38)}",
            87: previous_ref(39),
            90: statement_ref(39),
        }
        for row, value in expression.items():
            write_forecast(row, value)
    for column_number in historical_columns + forecast_columns:
        statement_ref = partial(ref, "Statements", column_number=column_number)
        for row, expression, dependencies in [
            (107, f"{statement_ref(36)}-{statement_ref(43)}", [36, 43]),
            (
                108,
                f"{statement_ref(17)}-{statement_ref(18)}-{statement_ref(19)}",
                [17, 18, 19],
            ),
            (
                109,
                f"{statement_ref(72)}+{statement_ref(62)}+{statement_ref(65)}+{statement_ref(69)}+{statement_ref(70)}-{statement_ref(73)}",
                [72, 62, 65, 69, 70, 73],
            ),
        ]:
            layout.write_formula(
                "Statements",
                row,
                column_number,
                expression,
                [statement_ref(row_number) for row_number in dependencies],
            )
            model_spec["checks"].append(
                {"cell": statement_ref(row), "label": labels[row], "tolerance": 0.5}
            )
        layout.write_cell(
            "Statements",
            110,
            column_number,
            f'=IF(ISNUMBER({statement_ref(31)}),IF({statement_ref(31)}<0,"Funding needed","Cash is non-negative"),"Inputs needed")',
            "General",
        )
    for index, (_, title, _number_format, _, _) in enumerate(DRIVERS):
        layout.write_label("Statements", 120 + index, title)
    for row_number, text in {
        9: "Previous sales × (1 + revenue growth).",
        12: "Gross profit less EBITDA. EBITDA itself is set by the margin input.",
        16: (
            "Actuals: EBIT less reported pre-tax profit. Forecast: interest on opening "
            "debt, less entered other income."
        ),
        35: (
            "Total assets less cash, receivables, inventory and fixed assets. This "
            "remainder stays constant in the forecast."
        ),
        40: "Total liabilities less trade payables and debt. This remainder stays constant in the forecast.",
        42: "Previous equity + net profit - dividends + net cash from shares issued or repurchased.",
        60: "Actuals reconcile reported operating cash. No additional noncash adjustments are forecast.",
        64: "Actuals: investing cash flow plus capital spending. No other investing cash flows are forecast.",
        66: "Actual debt cash movements are not inferred from changes in the balance sheet.",
        68: (
            "Actuals: financing cash flow, less net borrowing, plus dividends. Forecast: "
            "cash from share issues, less buybacks."
        ),
        70: "Actuals reconcile cash balances with reported cash flows. No separate FX cash movement is forecast.",
        86: "Receivables + inventory - trade payables.",
        107: "Zero means the balance sheet ties. Historical rounding tolerance is 0.5 in workbook units.",
    }.items():
        layout.write_cell("Statements", row_number, 14, text, "General", "explanation")
    layout.mark_totals("Statements", 11, 13, 15, 19, 36, 43, 62, 65, 69, 73)
    # A DCF built from the same forecast, not a second forecast model.
    layout.write_note(
        "DCF",
        4,
        (
            "Finish the forecast first, then enter the DCF inputs on Assumptions. Cash "
            "flows are discounted at each year end. A missing value stays blank; it does "
            "not mean zero."
        ),
    )
    for index, column_number in enumerate(range(4, 9)):
        layout.write_cell(
            "DCF",
            7,
            column_number,
            int(history[-1]["period"][:4]) + index + 1,
            '0" forecast"',
            "header",
        )
    dcf_labels = {
        9: "Operating profit (EBIT)",
        10: "Tax on operating profit",
        11: "Operating profit after tax",
        12: "Depreciation added back",
        13: "Capital spending (outflow)",
        14: "Increase in working capital",
        15: "Free cash flow to the firm",
        18: "Year-end discount factor",
        19: "Present value of cash flow",
        23: "Year-six cash flow",
        24: "Terminal value at year five",
        25: "Present value of terminal value",
        29: "Present value of five-year cash flows",
        30: "Present value of terminal value",
        31: "Operating business value",
        33: "Cash to add",
        34: "Debt to deduct",
        35: "Other claims to deduct",
        36: "Other non-operating assets to add",
        37: "Equity value",
        39: "Shares outstanding (millions)",
        40: "Value per share",
        42: "Current share price (optional)",
        43: "Upside / downside (optional)",
    }
    for row_number, text in dcf_labels.items():
        layout.write_label("DCF", row_number, text)
    wacc = ref("Assumptions", 31, 6)
    growth = ref("Assumptions", 32, 6)
    shares = ref("Assumptions", 33, 6)
    claims = ref("Assumptions", 34, 6)
    nonop = ref("Assumptions", 35, 6)
    price = ref("Assumptions", 36, 6)
    for index, (column_number, model_column) in enumerate(
        zip(range(4, 9), forecast_columns, strict=False), 1
    ):
        model_ref = partial(ref, "Statements", column_number=model_column)
        dcf_ref = partial(ref, "DCF", column_number=column_number)
        tax = ref("Assumptions", assumption_rows["tax_rate"], 5 + index)
        for row, model_row in [(9, 15), (12, 14), (13, 63), (14, 61)]:
            layout.write_formula(
                "DCF", row, column_number, model_ref(model_row), [model_ref(model_row)]
            )
        layout.write_formula(
            "DCF",
            10,
            column_number,
            f"MAX(0,{dcf_ref(9)})*{tax}",
            [dcf_ref(9), tax],
            f"AND({tax}>=0,{tax}<=1)",
        )
        layout.write_formula(
            "DCF",
            11,
            column_number,
            f"{dcf_ref(9)}-{dcf_ref(10)}",
            [dcf_ref(9), dcf_ref(10)],
        )
        layout.write_formula(
            "DCF",
            15,
            column_number,
            f"{dcf_ref(11)}+{dcf_ref(12)}-{dcf_ref(13)}-{dcf_ref(14)}",
            [dcf_ref(row_number) for row_number in (11, 12, 13, 14)],
        )
        layout.write_formula(
            "DCF",
            18,
            column_number,
            f"1/(1+{wacc})^{index}",
            [wacc],
            wacc + ">0",
            "0.000x",
        )
        layout.write_formula(
            "DCF",
            19,
            column_number,
            f"{dcf_ref(15)}*{dcf_ref(18)}",
            [dcf_ref(15), dcf_ref(18)],
        )
    dcf_ref = partial(ref, "DCF", column_number=4)
    layout.write_formula(
        "DCF", 23, 4, f"H15*(1+{growth})", ["H15", growth], growth + ">-1"
    )
    condition = f"AND({wacc}>0,{growth}>-1,{growth}<{wacc})"
    layout.write_formula(
        "DCF",
        24,
        4,
        f"{dcf_ref(23)}/({wacc}-{growth})",
        [dcf_ref(23), wacc, growth],
        condition,
    )
    layout.write_formula("DCF", 25, 4, f"{dcf_ref(24)}*H18", [dcf_ref(24), "H18"])
    layout.write_formula(
        "DCF",
        29,
        4,
        "SUM(D19:H19)",
        [f"{column_letter(column_number)}19" for column_number in range(4, 9)],
    )
    layout.write_formula("DCF", 30, 4, dcf_ref(25), [dcf_ref(25)])
    layout.write_formula(
        "DCF", 31, 4, f"{dcf_ref(29)}+{dcf_ref(30)}", [dcf_ref(29), dcf_ref(30)]
    )
    for row, target in [
        (33, ref("Statements", 31, last_actual_column)),
        (34, ref("Statements", 39, last_actual_column)),
        (35, claims),
        (36, nonop),
        (39, shares),
        (42, price),
    ]:
        layout.write_formula(
            "DCF",
            row,
            4,
            target,
            [target],
            target + ">=0"
            if row in (34, 35, 36)
            else target + ">0"
            if row in (39, 42)
            else "TRUE",
            PRICE if row == 42 else NUM,
        )
    layout.write_formula(
        "DCF",
        37,
        4,
        f"{dcf_ref(31)}+{dcf_ref(33)}-{dcf_ref(34)}-{dcf_ref(35)}+{dcf_ref(36)}",
        [dcf_ref(row_number) for row_number in (31, 33, 34, 35, 36)],
    )
    layout.write_formula(
        "DCF",
        40,
        4,
        f"{dcf_ref(37)}*{data['factor']}/({dcf_ref(39)}*1000000)",
        [dcf_ref(37), dcf_ref(39)],
        dcf_ref(39) + ">0",
        PRICE,
    )
    layout.write_formula(
        "DCF",
        43,
        4,
        f"{dcf_ref(40)}/{dcf_ref(42)}-1",
        [dcf_ref(40), dcf_ref(42)],
        dcf_ref(42) + ">0",
        PCT,
    )
    layout.write_cell("DCF", 6, 2, "Next step", "General", "label")
    layout.write_cell(
        "DCF",
        6,
        4,
        f'=IF(ISNUMBER(D40),"Review the result and model checks",IF(COUNT({wacc},{growth})<2,"Enter discount rate and long-run growth",IF({wacc}<=0,"Discount rate must be positive",IF(OR({growth}<=-1,{growth}>={wacc}),"Long-run growth must be below WACC",IF(COUNT(D15:H15)<5,"Complete the forecast on Assumptions","Enter shares, claims and non-operating assets")))))',
        "General",
        "status",
    )
    layout.worksheets["DCF"]["merges"].append("D6:H6")
    layout.add_section("DCF", 21, "Value after year five")
    layout.add_section("DCF", 27, "From business value to value per share")
    layout.mark_totals("DCF", 11, 15, 19, 31, 37, 40)
    layout.write_cell(
        "DCF",
        23,
        10,
        "Year-five cash flow × (1 + long-run growth). Check that future investment supports this growth.",
        "General",
        "explanation",
    )
    layout.write_cell(
        "DCF",
        24,
        10,
        "Year-six cash flow / (WACC - long-run growth). The growth rate must stay below WACC.",
        "General",
        "explanation",
    )
    layout.write_cell(
        "DCF",
        37,
        10,
        "Operating value + cash - debt - other claims + eligible non-operating assets.",
        "General",
        "explanation",
    )
    layout.write_note(
        "DCF",
        47,
        (
            "A negative cash flow or valuation is shown as calculated. Review whether the"
            " business and terminal-year assumptions are appropriate. The tax model gives"
            " no immediate credit for operating losses; it does not project tax-loss "
            "carryforwards."
        ),
    )
    layout.write_note("DCF", 51, "Method reference: " + METHOD)
    model_spec["outputs"] = {
        "Operating business value": dcf_ref(31),
        "Equity value": dcf_ref(37),
        "Value per share": dcf_ref(40),
    }
    model_spec["notes"] = {
        "source_url": data.get("source_url"),
        "retrieved": data.get("retrieved"),
        "history_notes": data.get("notes", []),
    }
    return model_spec
