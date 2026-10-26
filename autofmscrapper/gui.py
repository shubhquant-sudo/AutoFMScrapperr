"""A separate, small desktop front door for the beginner model."""

import json
import os
import queue
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .data import BASE, ticker_name
from .runner import generate, history, write_json
from .theme import BG, INK, SURFACE, configure


class AutoFMScrapperApp:
    def __init__(self, root):
        self.root = root
        self.busy = False
        self.closing = False
        self.events = queue.Queue()
        self.report = None
        self.import_file = None
        self.rows = []
        root.title("AutoFMScrapper")
        root.geometry("1020x780")
        root.minsize(860, 680)
        root.configure(bg=BG)
        configure(root)
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.report_callback_exception = self.callback_error
        prefs = BASE / "cache/preferences.json"
        try:
            saved = json.loads(prefs.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            saved = {}
        self.ticker = tk.StringVar(value=saved.get("ticker", ""))
        self.market = tk.StringVar(value=saved.get("market", "US"))
        self.refresh = tk.BooleanVar(value=False)
        self.source = tk.StringVar(
            value="Annual statements from Yahoo, or a selected Screener export."
        )
        self.status = tk.StringVar(
            value="Enter a ticker to start. Forecast and DCF inputs will be left blank."
        )
        body = self._build_page(root)
        card = self._build_company_form(body)
        self._build_generation_form(card)
        self._build_result_buttons(body)
        self._build_history_table(body)
        self.load_history()
        root.after(100, self.poll)

    def _build_page(self, root):
        self.canvas = tk.Canvas(root, bg=BG, highlightthickness=0)
        scroll = ttk.Scrollbar(root, command=self.canvas.yview)
        scroll.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=scroll.set)
        self.canvas.pack(fill="both", expand=True)
        body = ttk.Frame(self.canvas, padding=24)
        self.body = body
        item = self.canvas.create_window((0, 0), window=body, anchor="nw")
        self.canvas.bind(
            "<Configure>", lambda e: self.canvas.itemconfigure(item, width=e.width)
        )
        body.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        root.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        ttk.Label(body, text="AutoFMScrapper", style="Heading.TLabel").pack(anchor="w")
        ttk.Label(
            body,
            text="Download the company history, then enter your forecast assumptions in Excel.",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(7, 18))
        return body

    def _build_company_form(self, body):
        card = ttk.Frame(body, style="Panel.TFrame", padding=18)
        card.pack(fill="x")
        ttk.Label(card, text="1. Choose the company", style="CardTitle.TLabel").pack(
            anchor="w", pady=(0, 10)
        )
        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(fill="x")
        row.columnconfigure(0, weight=3)
        row.columnconfigure(1, weight=1)
        ttk.Label(row, text="Ticker", style="CardMuted.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 5)
        )
        ttk.Label(row, text="Market", style="CardMuted.TLabel").grid(
            row=0, column=1, sticky="w", pady=(0, 5)
        )
        values = {}
        for path in (BASE / "cache/snapshots").glob("*.json"):
            try:
                raw = json.loads(path.read_text(encoding="utf-8"))
                values[raw["ticker"]] = raw.get("company_name", raw["ticker"])
            except (OSError, ValueError, KeyError):
                continue
        self.ticker_box = ttk.Combobox(
            row,
            textvariable=self.ticker,
            values=[t + " · " + n for t, n in sorted(values.items())],
        )
        self.ticker_box.grid(row=1, column=0, sticky="ew", padx=(0, 12))
        self.market_box = ttk.Combobox(
            row,
            textvariable=self.market,
            values=["US", "India · NSE", "Ticker as entered"],
            state="readonly",
            width=18,
        )
        self.market_box.grid(row=1, column=1, sticky="ew")
        self.ticker_box.bind("<<ComboboxSelected>>", self.company_selected)
        self.ticker.trace_add("write", self.company_changed)
        controls = ttk.Frame(card, style="Card.TFrame")
        controls.pack(fill="x", pady=(12, 6))
        self.import_button = ttk.Button(
            controls, text="Import Screener export", command=self.choose_import
        )
        self.import_button.pack(side="left")
        self.clear_button = ttk.Button(
            controls, text="Clear import", command=self.clear_import
        )
        self.clear_button.pack(side="left", padx=8)
        self.source_label = ttk.Label(
            card, textvariable=self.source, style="CardMuted.TLabel", wraplength=850
        )
        self.source_label.pack(anchor="w", fill="x", pady=(2, 6))
        self.refresh_box = ttk.Checkbutton(
            card,
            text="Fetch new data instead of using the saved snapshot",
            variable=self.refresh,
        )
        self.refresh_box.pack(anchor="w")
        return card

    def _build_generation_form(self, card):
        ttk.Separator(card).pack(fill="x", pady=14)
        ttk.Label(card, text="2. Create your workbook", style="CardTitle.TLabel").pack(
            anchor="w", pady=(0, 7)
        )
        self.delivery_label = ttk.Label(
            card,
            text=(
                "Includes income statement, balance sheet, cash flow and DCF. All available "
                "history and original source files are retained. You fill the yellow forecast"
                " and valuation inputs."
            ),
            style="CardMuted.TLabel",
            wraplength=850,
        )
        self.delivery_label.pack(anchor="w", fill="x")
        self.create = ttk.Button(
            card,
            text="Create my workbook",
            style="Primary.TButton",
            command=self.create_workbook,
        )
        self.create.pack(anchor="w", pady=(12, 9))
        self.progress = ttk.Progressbar(card, maximum=5)
        self.progress.pack(fill="x", pady=(3, 8))
        self.status_label = ttk.Label(
            card, textvariable=self.status, style="CardMuted.TLabel", wraplength=850
        )
        self.status_label.pack(anchor="w", fill="x")
        card.bind("<Configure>", self._resize_form)
        self.controls = [
            self.ticker_box,
            self.market_box,
            self.import_button,
            self.clear_button,
            self.refresh_box,
            self.create,
        ]

    def _build_result_buttons(self, body):
        bar = ttk.Frame(body)
        bar.pack(fill="x", pady=(14, 14))
        self.open_button = ttk.Button(
            bar, text="Open workbook", command=lambda: self.open_result("workbook")
        )
        self.open_button.pack(side="left")
        self.folder_button = ttk.Button(
            bar, text="Workbook and sources", command=lambda: self.open_result("folder")
        )
        self.folder_button.pack(side="left", padx=8)
        self.zip_button = ttk.Button(
            bar, text="Complete ZIP", command=lambda: self.open_result("zip")
        )
        self.zip_button.pack(side="left")
        ttk.Button(bar, text="How to use it", command=self.show_help).pack(side="right")
        for button in [self.open_button, self.folder_button, self.zip_button]:
            button.state(["disabled"])

    def _build_history_table(self, body):
        ttk.Label(body, text="Recent workbooks", style="CardTitle.TLabel").pack(
            anchor="w", pady=(0, 7)
        )
        table = ttk.Frame(body)
        table.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(
            table,
            columns=("ticker", "periods", "date", "status"),
            show="headings",
            height=3,
        )
        for key, title, width in [
            ("ticker", "TICKER", 110),
            ("periods", "HISTORY", 110),
            ("date", "CREATED (UTC)", 180),
            ("status", "STATUS", 200),
        ]:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, minwidth=70)
        scroll = ttk.Scrollbar(table, command=self.tree.yview)
        scroll.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.select_history)
        self.tree.bind("<Double-1>", lambda e: self.open_result("workbook"))

    def company_selected(self, event=None):
        symbol = self.ticker.get().split(" · ", 1)[0]
        self.market.set("India · NSE" if symbol.endswith((".NS", ".BO")) else "US")

    def _on_mousewheel(self, event):
        if str(event.widget).startswith(str(self.body)) and not isinstance(
            event.widget, ttk.Combobox
        ):
            self.canvas.yview_scroll(-int(event.delta / 120), "units")

    def _resize_form(self, event):
        width = max(200, event.width - 40)
        for label in (self.status_label, self.source_label, self.delivery_label):
            label.configure(wraplength=width)

    def company_changed(self, *_):
        if self.import_file and self.ticker.get() != getattr(
            self, "import_company", ""
        ):
            self.clear_import()

    def choose_import(self):
        path = filedialog.askopenfilename(
            parent=self.root,
            title="Choose a Screener company export",
            initialdir=str(BASE / "imports"),
            filetypes=[("Company export", "*.xlsx *.csv")],
        )
        if path:
            self.import_file = Path(path)
            self.import_company = self.ticker.get()
            self.source.set("Selected export: " + self.import_file.name)

    def clear_import(self):
        self.import_file = None
        self.source.set("Annual statements from Yahoo, or a selected Screener export.")

    def request(self):
        ticker = self.ticker.get().split(" · ", 1)[0].strip().upper()
        if (
            self.market.get() == "India · NSE"
            and not ticker.endswith((".NS", ".BO"))
            and ticker
        ):
            ticker += ".NS"
        ticker = ticker_name(ticker)
        if self.import_file:
            if not ticker.endswith((".NS", ".BO")):
                raise ValueError(
                    "Use an Indian company ticker for a Screener export, or clear the import."
                )
            if not self.import_file.is_file():
                raise ValueError("The selected export is no longer available.")
        return {
            "ticker": ticker,
            "import_file": self.import_file,
            "refresh": self.refresh.get(),
        }

    def set_busy(self, value):
        self.busy = value
        for widget in self.controls:
            widget.state(["disabled"] if value else ["!disabled"])
        if not value:
            self.market_box.state(["readonly"])

    def create_workbook(self):
        if self.busy:
            return
        try:
            request = self.request()
        except ValueError as error:
            self.status.set(str(error))
            return
        write_json(
            BASE / "cache/preferences.json",
            {"ticker": self.ticker.get(), "market": self.market.get()},
        )
        self.set_busy(True)
        self.refresh.set(False)
        self.progress["value"] = 0
        self.status.set("Starting…")
        threading.Thread(target=self.worker, args=(request,), daemon=True).start()

    def worker(self, request):
        try:
            self.events.put(
                (
                    "done",
                    generate(
                        **request,
                        progress=lambda n, t: self.events.put(("progress", n, t)),
                    ),
                )
            )
        except Exception as error:
            self.events.put(("error", str(error), traceback.format_exc()))

    def handle(self, item):
        if item[0] == "progress":
            self.progress["value"] = item[1]
            self.status.set(item[2])
        elif item[0] == "done":
            self.set_busy(False)
            self.report = item[1]
            self.progress["value"] = 5
            detail = f"Saved {len(self.report['history_periods'])} years of history. Forecast and DCF inputs are blank."
            if self.report["missing_latest_source_fields"]:
                detail += (
                    " Check the yellow History cells for: "
                    + ", ".join(self.report["missing_latest_source_fields"])
                    + "."
                )
            if self.report.get("historical_checks_to_review"):
                detail += " Review the historical differences under Statements checks."
            if self.report.get("opening_source_notes"):
                detail += " " + " ".join(self.report["opening_source_notes"])
            self.status.set(detail)
            for button in [self.open_button, self.folder_button, self.zip_button]:
                button.state(["!disabled"])
            self.load_history()
        elif item[0] == "error":
            self.set_busy(False)
            self.status.set(item[1])
            log = BASE / "cache/errors.log"
            log.parent.mkdir(parents=True, exist_ok=True)
            with log.open("a", encoding="utf-8") as file:
                file.write(item[2] + "\n")
        if item[0] in ("done", "error") and self.closing:
            self.root.destroy()
            return True
        return False

    def poll(self):
        try:
            while True:
                if self.handle(self.events.get_nowait()):
                    return
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def load_history(self):
        self.rows = history()
        self.tree.delete(*self.tree.get_children())
        for i, report in enumerate(self.rows):
            self.tree.insert(
                "",
                "end",
                iid=str(i),
                values=(
                    report["ticker"],
                    str(len(report["history_periods"])) + " years",
                    report["created_utc"].replace("T", " ")[:16],
                    "Check history"
                    if report["missing_latest_source_fields"]
                    or report.get("historical_checks_to_review")
                    or report.get("opening_source_notes")
                    else "Blank inputs ready",
                ),
            )

    def select_history(self, event=None):
        selected = self.tree.selection()
        if selected:
            self.report = self.rows[int(selected[0])]
            for button in [self.open_button, self.folder_button, self.zip_button]:
                button.state(["!disabled"])

    def open_path(self, path):
        try:
            if not Path(path).exists():
                raise OSError("This file or folder is no longer available.")
            os.startfile(str(Path(path).resolve()))
        except OSError as error:
            messagebox.showerror("Could not open", str(error), parent=self.root)

    def open_result(self, key):
        if self.report:
            self.open_path(self.report[key])

    def show_help(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Using AutoFMScrapper")
        dialog.geometry("800x650")
        dialog.minsize(600, 450)
        frame = ttk.Frame(dialog, padding=18)
        frame.pack(fill="both", expand=True)
        text = tk.Text(
            frame,
            wrap="word",
            bg=SURFACE,
            fg=INK,
            font=("Segoe UI", 10),
            relief="flat",
            padx=18,
            pady=18,
        )
        scroll = ttk.Scrollbar(frame, command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        text.pack(fill="both", expand=True)
        text.insert("1.0", (BASE / "README.md").read_text(encoding="utf-8"))
        text.configure(state="disabled")

    def callback_error(self, kind, error, tb):
        messagebox.showerror("Please check this step", str(error), parent=self.root)

    def close(self):
        if self.busy:
            self.closing = messagebox.askyesno(
                "Finish saving first?",
                "The workbook is being prepared. Close this window after the current run finishes?",
                parent=self.root,
            )
        else:
            self.root.destroy()


def main():
    root = tk.Tk()
    AutoFMScrapperApp(root)
    root.mainloop()
