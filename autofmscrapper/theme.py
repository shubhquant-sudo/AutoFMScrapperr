"""Shared dark Studio palette and ttk styling."""

from tkinter import ttk

BG = "#171B21"
NAV = "#101419"
SURFACE = "#222933"
INK = "#EDF2F7"
MUTED = "#A4B0C0"
ACCENT = "#62CABB"
LINE = "#35414F"
SELECT = "#294A4D"


def configure(root):
    root.option_add("*Font", ("Segoe UI", 10))
    for key, value in (
        ("*TCombobox*Listbox.background", SURFACE),
        ("*TCombobox*Listbox.foreground", INK),
        ("*TCombobox*Listbox.selectBackground", SELECT),
        ("*TCombobox*Listbox.selectForeground", INK),
    ):
        root.option_add(key, value)
    s = ttk.Style(root)
    s.theme_use("clam")
    for name, color in (("TFrame", BG), ("Card.TFrame", SURFACE)):
        s.configure(name, background=color)
    s.configure(
        "Panel.TFrame",
        background=SURFACE,
        borderwidth=1,
        relief="solid",
        bordercolor=LINE,
    )
    s.configure("TLabelframe", background=SURFACE, bordercolor=LINE)
    s.configure("TLabelframe.Label", background=SURFACE, foreground=MUTED)
    for name, bg, fg in (
        ("TLabel", BG, INK),
        ("Card.TLabel", SURFACE, INK),
        ("Muted.TLabel", BG, MUTED),
        ("CardMuted.TLabel", SURFACE, MUTED),
    ):
        s.configure(name, background=bg, foreground=fg)
    s.configure("Heading.TLabel", font=("Segoe UI", 24, "bold"), foreground=INK)
    s.configure("Kicker.TLabel", font=("Segoe UI", 9, "bold"), foreground=ACCENT)
    s.configure(
        "CardTitle.TLabel",
        background=SURFACE,
        font=("Segoe UI", 12, "bold"),
        foreground=INK,
    )
    s.configure(
        "Badge.TLabel",
        background=SELECT,
        foreground=ACCENT,
        padding=(10, 5),
        font=("Segoe UI", 9, "bold"),
    )
    s.configure(
        "TButton",
        padding=(13, 9),
        background="#303B48",
        foreground=INK,
        borderwidth=0,
        focusthickness=2,
        focuscolor=ACCENT,
    )
    s.map(
        "TButton",
        background=[("active", "#405060")],
        foreground=[("disabled", "#788696")],
    )
    s.configure(
        "Primary.TButton",
        background=ACCENT,
        foreground=NAV,
        font=("Segoe UI", 10, "bold"),
        padding=(17, 11),
    )
    s.map(
        "Primary.TButton",
        background=[("disabled", "#354E50"), ("active", "#87DED1")],
        foreground=[("disabled", "#A0B2B3")],
    )
    s.configure("Quiet.TButton", background=SURFACE, foreground=MUTED, padding=(9, 7))
    s.map("Quiet.TButton", background=[("active", "#303B48")])
    for name in ("TEntry", "TCombobox"):
        s.configure(
            name,
            padding=(10, 8),
            fieldbackground=NAV,
            background=SURFACE,
            foreground=INK,
            bordercolor=LINE,
            lightcolor=LINE,
            darkcolor=LINE,
            insertcolor=INK,
            arrowsize=13,
        )
        s.map(
            name,
            fieldbackground=[("disabled", SURFACE), ("readonly", NAV)],
            foreground=[("disabled", "#788696"), ("readonly", INK)],
            bordercolor=[("focus", ACCENT)],
        )
    for name, bg in (
        ("TCheckbutton", SURFACE),
        ("TRadiobutton", SURFACE),
        ("Page.TCheckbutton", BG),
    ):
        s.configure(
            name, background=bg, foreground=INK, padding=(0, 5), indicatormargin=5
        )
        s.map(
            name,
            background=[("active", bg)],
            foreground=[("disabled", "#788696")],
            indicatorbackground=[("selected", ACCENT), ("!selected", NAV)],
            indicatorforeground=[("selected", NAV)],
        )
    s.configure(
        "Treeview",
        rowheight=39,
        font=("Segoe UI", 10),
        background=SURFACE,
        fieldbackground=SURFACE,
        foreground=INK,
        borderwidth=0,
    )
    s.configure(
        "Treeview.Heading",
        font=("Segoe UI", 9, "bold"),
        background="#303B48",
        foreground=MUTED,
        padding=(9, 10),
        relief="flat",
    )
    s.map("Treeview", background=[("selected", SELECT)], foreground=[("selected", INK)])
    s.configure(
        "Horizontal.TProgressbar",
        background=ACCENT,
        troughcolor=NAV,
        borderwidth=0,
        thickness=4,
    )
    s.configure("TSeparator", background=LINE)
    s.configure(
        "TScrollbar",
        background="#3B4857",
        troughcolor=BG,
        bordercolor=BG,
        arrowcolor=MUTED,
    )
    return s
