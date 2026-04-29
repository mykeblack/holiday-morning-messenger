#!/usr/bin/env python3
"""
Holiday Morning Messenger  v2
3-step wizard: greeting → occasion → link → copy to Slack
API sources are defined in config.json.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import re
import threading
import webbrowser
import platform
import subprocess
import sys
from datetime import datetime
from urllib.request import urlopen, Request
from urllib.parse import quote_plus
from urllib.error import URLError, HTTPError

# ─────────────────────────────────────────────────────────────────────────────
#  PATHS
# ─────────────────────────────────────────────────────────────────────────────
APP_DIR     = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "config.json")

# ─────────────────────────────────────────────────────────────────────────────
#  DEFAULT CONFIG  (written to config.json on first run)
# ─────────────────────────────────────────────────────────────────────────────
DEFAULT_CONFIG = {
    "project_name": "Team",
    "sensitivity_filter": True,
    "greetings": [
        "Good morning everyone",
        "Morning all",
        "Morning {project} people",
        "Hey team",
        "Rise and shine, team",
        "Happy {day}, everyone"
    ],
    "recent_greetings": [],
    "apis": [
        {
            "id":       "wikipedia",
            "name":     "Wikipedia – Observances & Holidays",
            "type":     "wikipedia",
            "language": "en",
            "enabled":  True,
            "notes":    "No API key required. Pulls observances from Wikipedia date pages."
        },
        {
            "id":      "nager",
            "name":    "Nager.Date – Public Holidays",
            "type":    "nager",
            "country": "GB",
            "enabled": True,
            "notes":   "No API key required. Change 'country' to your ISO country code."
        },
        {
            "id":       "calendarific",
            "name":     "Calendarific",
            "type":     "calendarific",
            "api_key":  "",
            "country":  "GB",
            "enabled":  False,
            "notes":    "Free API key at https://calendarific.com – supports 230+ countries."
        },
        {
            "id":      "abstractapi",
            "name":    "Abstract API – Holidays",
            "type":    "abstractapi",
            "api_key": "",
            "country": "GB",
            "enabled": False,
            "notes":   "Free API key at https://app.abstractapi.com/api/holidays"
        }
    ]
}

# ─────────────────────────────────────────────────────────────────────────────
#  SENSITIVITY FILTER
# ─────────────────────────────────────────────────────────────────────────────
# Hard block – saying "Happy X" is clearly inappropriate
_HARD_BLOCK = [
    "holocaust", "genocide", "hiroshima", "nagasaki", "nakba",
    "day of mourning", "kristallnacht", "srebrenica",
    "victims of aggression", "rwandan genocide", "armenian genocide",
    "day of memory and sorrow", "victims of fascism",
    "day of national tragedy", "black ribbon day", "remembrance of victims",
]

# Soft flag – worth a ⚠️ but user can still choose
_SOFT_WARN = [
    "remembrance", "memorial", "fallen", "cancer", "hiv", "aids",
    "human trafficking", "slavery", "violence", "abuse",
    "missing", "poverty", "hunger", "homelessness", "exploitation",
    "epidemic", "pandemic", "drug abuse", "leprosy",
]

def sensitivity(name: str) -> str:
    """Returns 'block', 'warn', or 'ok'."""
    lo = name.lower()
    if any(k in lo for k in _HARD_BLOCK):
        return "block"
    if any(k in lo for k in _SOFT_WARN):
        return "warn"
    return "ok"

# Words that suggest a name is a person / location, not an observance
_NOT_AN_OBSERVANCE = re.compile(
    r"^(january|february|march|april|may|june|july|august|september|"
    r"october|november|december|\d{4}|saint |st\.|feast of)",
    re.I
)

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def load_config() -> dict:
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                cfg = json.load(f)
            for k, v in DEFAULT_CONFIG.items():
                if k not in cfg:
                    cfg[k] = v
            return cfg
        except Exception:
            pass
    cfg = {k: v for k, v in DEFAULT_CONFIG.items()}
    save_config(cfg)
    return cfg

def save_config(cfg: dict):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)

# ─────────────────────────────────────────────────────────────────────────────
#  HTTP HELPER
# ─────────────────────────────────────────────────────────────────────────────
_UA = {"User-Agent": "HolidayMessenger/2.0 (https://github.com/yourname/holiday-messenger)"}

def fetch_json(url: str, timeout: int = 9):
    try:
        req = Request(url, headers=_UA)
        with urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception:
        return None

# ─────────────────────────────────────────────────────────────────────────────
#  API FETCHERS
# ─────────────────────────────────────────────────────────────────────────────
MONTH_NAMES = ["January","February","March","April","May","June",
               "July","August","September","October","November","December"]

def _holiday(name, source, url="", sens=None) -> dict:
    if sens is None:
        sens = sensitivity(name)
    return {"name": name, "source": source, "url": url, "sensitivity": sens}

# ── Wikipedia ─────────────────────────────────────────────────────────────────
def fetch_wikipedia(date: datetime, language: str = "en") -> list[dict]:
    page = f"{MONTH_NAMES[date.month-1]}_{date.day}"
    base = f"https://{language}.wikipedia.org/w/api.php"

    # Step 1: get sections to find "Holidays and observances"
    secs = fetch_json(f"{base}?action=parse&page={page}&prop=sections&format=json")
    if not secs:
        return []
    sections = (secs.get("parse") or {}).get("sections") or []
    section_idx = None
    for s in sections:
        title_lo = (s.get("line") or "").lower()
        if "observance" in title_lo or ("holiday" in title_lo and "event" not in title_lo):
            section_idx = s.get("index")
            break
    if not section_idx:
        return []

    # Step 2: get wikitext of that section
    content = fetch_json(
        f"{base}?action=parse&page={page}&prop=wikitext&section={section_idx}&format=json"
    )
    if not content:
        return []
    wikitext = ((content.get("parse") or {}).get("wikitext") or {}).get("*","")

    results = []
    seen    = set()
    for line in wikitext.split("\n"):
        line = line.strip()
        if not line.startswith("*"):
            continue
        # Extract first [[Target|Display]] or [[Target]]
        m = re.search(r"\[\[([^\]|#]+?)(?:\|[^\]]+)?\]\]", line)
        if not m:
            continue
        name = m.group(1).strip()
        if name in seen:
            continue
        if re.match(r"^\d{4}$", name):
            continue
        if _NOT_AN_OBSERVANCE.match(name):
            continue
        # Skip very short strings (likely abbreviations / months)
        if len(name) < 5:
            continue
        sens = sensitivity(name)
        if sens == "block":
            continue
        seen.add(name)
        wiki_url = f"https://{language}.wikipedia.org/wiki/{name.replace(' ','_')}"
        results.append(_holiday(name, "Wikipedia", wiki_url, sens))

    return results

# ── Nager.Date ────────────────────────────────────────────────────────────────
def fetch_nager(date: datetime, country: str = "GB") -> list[dict]:
    data = fetch_json(f"https://date.nager.at/api/v3/PublicHolidays/{date.year}/{country}")
    if not data:
        return []
    target = date.strftime("%Y-%m-%d")
    results = []
    for h in data:
        if h.get("date") == target:
            name = h.get("name") or h.get("localName") or ""
            if name:
                sens = sensitivity(name)
                if sens != "block":
                    results.append(_holiday(name, "Nager.Date", "", sens))
    return results

# ── Calendarific ──────────────────────────────────────────────────────────────
def fetch_calendarific(date: datetime, api_key: str, country: str = "GB") -> list[dict]:
    if not api_key.strip():
        return []
    url = (
        f"https://calendarific.com/api/v2/holidays"
        f"?api_key={api_key}&country={country}"
        f"&year={date.year}&month={date.month}&day={date.day}"
    )
    data = fetch_json(url)
    if not data:
        return []
    hols = ((data.get("response") or {}).get("holidays")) or []
    results = []
    for h in hols:
        name = h.get("name","")
        if name:
            sens = sensitivity(name)
            if sens != "block":
                results.append(_holiday(name, "Calendarific", "", sens))
    return results

# ── Abstract API ──────────────────────────────────────────────────────────────
def fetch_abstractapi(date: datetime, api_key: str, country: str = "GB") -> list[dict]:
    if not api_key.strip():
        return []
    url = (
        f"https://holidays.abstractapi.com/v1/"
        f"?api_key={api_key}&country={country}"
        f"&year={date.year}&month={date.month}&day={date.day}"
    )
    data = fetch_json(url)
    if not isinstance(data, list):
        return []
    results = []
    for h in data:
        name = h.get("name","")
        if name:
            sens = sensitivity(name)
            if sens != "block":
                results.append(_holiday(name, "Abstract API", "", sens))
    return results

# ── Coordinator ───────────────────────────────────────────────────────────────
def fetch_all_holidays(cfg: dict, date: datetime) -> list[dict]:
    """Run all enabled APIs and return a deduplicated, sorted list."""
    results = []
    for api in cfg.get("apis", []):
        if not api.get("enabled"):
            continue
        try:
            t = api["type"]
            if t == "wikipedia":
                results += fetch_wikipedia(date, api.get("language","en"))
            elif t == "nager":
                results += fetch_nager(date, api.get("country","GB"))
            elif t == "calendarific":
                results += fetch_calendarific(date, api.get("api_key",""), api.get("country","GB"))
            elif t == "abstractapi":
                results += fetch_abstractapi(date, api.get("api_key",""), api.get("country","GB"))
        except Exception:
            pass

    # Deduplicate (case-insensitive name match)
    seen, unique = set(), []
    for h in results:
        key = h["name"].lower()
        if key not in seen:
            seen.add(key)
            unique.append(h)

    # Sort: ok → warn (blocks already removed)
    unique.sort(key=lambda h: (0 if h["sensitivity"] == "ok" else 1, h["name"]))
    return unique

# ─────────────────────────────────────────────────────────────────────────────
#  LINK DISCOVERY
# ─────────────────────────────────────────────────────────────────────────────
def find_links(holiday: dict) -> list[dict]:
    """Return up to 6 {label, url, icon, type} dicts for a chosen holiday."""
    name  = holiday["name"]
    links = []
    seen  = set()

    def add(label, url, icon, ltype):
        if url and url not in seen:
            seen.add(url)
            links.append({"label": label, "url": url, "icon": icon, "type": ltype})

    # 1. Wikipedia search
    q    = quote_plus(name)
    data = fetch_json(
        f"https://en.wikipedia.org/w/api.php"
        f"?action=opensearch&search={q}&limit=3&namespace=0&format=json"
    )
    if data and len(data) >= 4:
        for title, url in zip(data[1], data[3]):
            add(f"Wikipedia: {title}", url, "📖", "wikipedia")

    # 2. Source URL (e.g. if holiday came from Wikipedia it already has one)
    if holiday.get("url"):
        src_url = holiday["url"]
        if src_url not in seen:
            add(f"{holiday['source']} page", src_url, "🔗", "source")

    # 3. DuckDuckGo Instant Answer
    ddg = fetch_json(
        f"https://api.duckduckgo.com/?q={q}&format=json&no_redirect=1&no_html=1&skip_disambig=1"
    )
    if ddg:
        if ddg.get("AbstractURL"):
            src = ddg.get("AbstractSource") or "Info"
            add(src, ddg["AbstractURL"], "🌐", "official")
        if ddg.get("OfficialSite"):
            add("Official website", ddg["OfficialSite"], "🌐", "official")
        for topic in (ddg.get("RelatedTopics") or [])[:4]:
            if isinstance(topic, dict) and topic.get("FirstURL"):
                text = re.sub(r"<[^>]+>", "", topic.get("Text") or "")[:65]
                add(text or "Related link", topic["FirstURL"], "💡", "related")

    # 4. UN / WHO pattern-matched URLs (for common international observances)
    slug = name.lower().replace("'","").replace(",","")
    slug = re.sub(r"[^a-z0-9]+", "-", slug).strip("-")
    un_url  = f"https://www.un.org/en/observances/{slug}"
    who_url = f"https://www.who.int/campaigns/{slug}"
    # Only add if looks like an international day
    if any(w in name.lower() for w in ["world ", "international ", "global "]):
        add("UN Observance page", un_url, "🇺🇳", "official")
        if any(w in name.lower() for w in ["health", "disease", "cancer", "aids", "diabetes",
                                            "blood", "nurse", "tobacco", "patient"]):
            add("WHO Campaign page", who_url, "🏥", "official")

    return links[:7]

# ─────────────────────────────────────────────────────────────────────────────
#  THEME
# ─────────────────────────────────────────────────────────────────────────────
BG      = "#FFFFFF"
SURFACE = "#F6F6F8"
BORDER  = "#E0E0E0"
ACCENT  = "#4F46E5"
ACCD    = "#3730A3"
PRI     = "#111111"
SEC     = "#555555"
MUT     = "#999999"
WARN_BG = "#FFFBEB"
WARN_FG = "#92400E"
GREEN   = "#16A34A"
RED     = "#DC2626"

FT  = ("Helvetica Neue", 13, "bold")   # title
FB  = ("Helvetica Neue", 12)           # body
FS  = ("Helvetica Neue", 10)           # small
FM  = ("Helvetica Neue", 12)           # message text
FC  = ("Courier", 11)                  # code/url

def _btn(parent, text, cmd, bg, fg, abg, afg, px=12, py=5, **kw):
    return tk.Button(parent, text=text, command=cmd, font=FB,
                     bg=bg, fg=fg, activebackground=abg, activeforeground=afg,
                     relief="flat", padx=px, pady=py, cursor="hand2", bd=0, **kw)

def primary_btn(parent, text, cmd, **kw):
    return _btn(parent, text, cmd, ACCENT, "#FFF", ACCD, "#FFF", **kw)

def ghost_btn(parent, text, cmd, **kw):
    return _btn(parent, text, cmd, SURFACE, SEC, BORDER, PRI, px=10, py=4, font=FS, **kw)

def danger_btn(parent, text, cmd, **kw):
    return _btn(parent, text, cmd, "#FEF2F2", RED, "#FEE2E2", RED, px=10, py=4, font=FS, **kw)

def hline(parent, **kw):
    tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", **kw)

# ─────────────────────────────────────────────────────────────────────────────
#  SETTINGS WINDOW
# ─────────────────────────────────────────────────────────────────────────────
class SettingsWindow(tk.Toplevel):
    def __init__(self, parent, cfg: dict, on_save):
        super().__init__(parent)
        self.title("Settings")
        self.cfg      = cfg
        self.on_save  = on_save
        self.configure(bg=BG)
        self.resizable(False, False)
        self.grab_set()
        self._build()

    def _build(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=0, pady=0)

        self._tab_general(nb)
        self._tab_apis(nb)
        self._tab_greetings(nb)
        self._tab_startup(nb)

        hline(self)
        footer = tk.Frame(self, bg=BG)
        footer.pack(fill="x", padx=16, pady=10)
        primary_btn(footer, "Save & close", self._save).pack(side="right", padx=(8,0))
        ghost_btn(footer, "Cancel", self.destroy).pack(side="right")

    # ── General tab ──────────────────────────────────────────────────────────
    def _tab_general(self, nb):
        frm = tk.Frame(nb, bg=BG, padx=20, pady=16)
        nb.add(frm, text="General")

        tk.Label(frm, text="Project / team name", font=FB, bg=BG, fg=PRI).pack(anchor="w")
        self.proj_var = tk.StringVar(value=self.cfg.get("project_name","Team"))
        tk.Entry(frm, textvariable=self.proj_var, font=FB, width=32,
                 relief="solid", bd=1).pack(anchor="w", pady=(4,14))

        self.filter_var = tk.BooleanVar(value=self.cfg.get("sensitivity_filter", True))
        tk.Checkbutton(frm, text="Filter out culturally insensitive days",
                       variable=self.filter_var, font=FB, bg=BG, fg=PRI,
                       activebackground=BG, selectcolor=BG).pack(anchor="w")
        tk.Label(frm,
                 text="When enabled, days like Holocaust Remembrance Day\n"
                      "are hidden and others are flagged with ⚠️.",
                 font=FS, bg=BG, fg=MUT, justify="left").pack(anchor="w", pady=(2,0))

    # ── APIs tab ──────────────────────────────────────────────────────────────
    def _tab_apis(self, nb):
        outer = tk.Frame(nb, bg=BG)
        nb.add(outer, text="API Sources")

        canvas = tk.Canvas(outer, bg=BG, bd=0, highlightthickness=0)
        sb     = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        frm    = tk.Frame(canvas, bg=BG, padx=20, pady=16)
        frm.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=frm, anchor="nw")
        canvas.configure(yscrollcommand=sb.set, height=340, width=460)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self._api_widgets = []
        for api in self.cfg.get("apis", []):
            self._api_card(frm, api)

    def _api_card(self, parent, api):
        card = tk.Frame(parent, bg=SURFACE, relief="flat",
                        padx=12, pady=10)
        card.pack(fill="x", pady=6)
        tk.Frame(card, bg=BORDER, height=1).pack(fill="x", side="bottom")

        top = tk.Frame(card, bg=SURFACE)
        top.pack(fill="x")
        ev = tk.BooleanVar(value=api.get("enabled", False))
        cb = tk.Checkbutton(top, text=api["name"], variable=ev, font=FB,
                            bg=SURFACE, fg=PRI, activebackground=SURFACE,
                            selectcolor=SURFACE)
        cb.pack(side="left")

        notes = api.get("notes","")
        if notes:
            tk.Label(card, text=notes, font=FS, bg=SURFACE,
                     fg=MUT, wraplength=400, justify="left").pack(anchor="w", pady=(4,0))

        widgets = {"enabled_var": ev, "api": api}

        if api["type"] in ("calendarific","abstractapi"):
            row = tk.Frame(card, bg=SURFACE)
            row.pack(fill="x", pady=(6,0))
            tk.Label(row, text="API key:", font=FS, bg=SURFACE, fg=SEC).pack(side="left")
            kv = tk.StringVar(value=api.get("api_key",""))
            tk.Entry(row, textvariable=kv, font=FC, width=36,
                     relief="solid", bd=1, show="*").pack(side="left", padx=(6,0))
            widgets["key_var"] = kv

        if api["type"] in ("nager","calendarific","abstractapi"):
            row2 = tk.Frame(card, bg=SURFACE)
            row2.pack(fill="x", pady=(4,0))
            tk.Label(row2, text="Country ISO code:", font=FS,
                     bg=SURFACE, fg=SEC).pack(side="left")
            cv = tk.StringVar(value=api.get("country","GB"))
            tk.Entry(row2, textvariable=cv, font=FB, width=6,
                     relief="solid", bd=1).pack(side="left", padx=(6,0))
            widgets["country_var"] = cv

        self._api_widgets.append(widgets)

    # ── Greetings tab ─────────────────────────────────────────────────────────
    def _tab_greetings(self, nb):
        frm = tk.Frame(nb, bg=BG, padx=20, pady=16)
        nb.add(frm, text="Greetings")

        tk.Label(frm, text="Greeting templates", font=FT, bg=BG, fg=PRI).pack(anchor="w")
        tk.Label(frm, text="Use {project} and {day} as placeholders.",
                 font=FS, bg=BG, fg=MUT).pack(anchor="w", pady=(2,10))

        list_frm = tk.Frame(frm, bg=BG, relief="solid", bd=1)
        list_frm.pack(fill="x")
        sb = tk.Scrollbar(list_frm)
        sb.pack(side="right", fill="y")
        self.greet_lb = tk.Listbox(list_frm, font=FB, yscrollcommand=sb.set,
                                   bg=BG, fg=PRI, selectbackground=ACCENT,
                                   selectforeground="#FFF", relief="flat",
                                   height=8, activestyle="none")
        self.greet_lb.pack(side="left", fill="both", expand=True)
        sb.config(command=self.greet_lb.yview)
        for g in self.cfg.get("greetings",[]):
            self.greet_lb.insert("end", g)

        btns = tk.Frame(frm, bg=BG)
        btns.pack(fill="x", pady=(8,0))
        ghost_btn(btns, "＋ Add",    self._add_greeting).pack(side="left", padx=(0,4))
        ghost_btn(btns, "✎ Edit",   self._edit_greeting).pack(side="left", padx=(0,4))
        danger_btn(btns,"✕ Delete", self._del_greeting).pack(side="left")

    def _add_greeting(self):
        self._greeting_dialog("", is_new=True)

    def _edit_greeting(self):
        sel = self.greet_lb.curselection()
        if not sel:
            messagebox.showinfo("Edit", "Select a greeting first.", parent=self)
            return
        self._greeting_dialog(self.greet_lb.get(sel[0]), is_new=False, idx=sel[0])

    def _del_greeting(self):
        sel = self.greet_lb.curselection()
        if not sel:
            return
        if self.greet_lb.size() <= 1:
            messagebox.showwarning("Delete", "Keep at least one greeting.", parent=self)
            return
        if messagebox.askyesno("Delete greeting", "Delete this greeting?", parent=self):
            self.greet_lb.delete(sel[0])

    def _greeting_dialog(self, text, is_new, idx=None):
        win = tk.Toplevel(self)
        win.title("Add greeting" if is_new else "Edit greeting")
        win.configure(bg=BG)
        win.resizable(False, False)
        win.grab_set()

        tk.Label(win, text="Greeting text:", font=FB, bg=BG, fg=PRI).pack(
            anchor="w", padx=16, pady=(12,4))
        var = tk.StringVar(value=text)
        tk.Entry(win, textvariable=var, font=FB, width=42,
                 relief="solid", bd=1).pack(padx=16)
        tk.Label(win, text="Placeholders: {project}  {day}", font=FS,
                 bg=BG, fg=MUT).pack(anchor="w", padx=16, pady=(4,0))

        def save():
            val = var.get().strip()
            if not val:
                return
            if is_new:
                self.greet_lb.insert("end", val)
            else:
                self.greet_lb.delete(idx)
                self.greet_lb.insert(idx, val)
            win.destroy()

        row = tk.Frame(win, bg=BG)
        row.pack(fill="x", padx=16, pady=12)
        primary_btn(row, "Save", save).pack(side="right", padx=(8,0))
        ghost_btn(row, "Cancel", win.destroy).pack(side="right")

    # ── Startup tab ───────────────────────────────────────────────────────────
    def _tab_startup(self, nb):
        frm = tk.Frame(nb, bg=BG, padx=20, pady=16)
        nb.add(frm, text="Startup")

        sys_name = platform.system()
        tk.Label(frm, text=f"Detected OS: {sys_name}", font=FB,
                 bg=BG, fg=PRI).pack(anchor="w", pady=(0,12))

        if sys_name == "Windows":
            tk.Label(frm,
                     text="Click below to add a launcher to your Windows\n"
                          "Startup folder. The app will open each time you log in.",
                     font=FB, bg=BG, fg=SEC, justify="left").pack(anchor="w")
            primary_btn(frm, "⚙  Install startup launcher (Windows)",
                        self._setup_windows).pack(anchor="w", pady=(12,0))
            ghost_btn(frm, "Open Startup folder in Explorer",
                      lambda: subprocess.Popen(
                          f'explorer /select,"{_startup_bat_path()}"'
                      )).pack(anchor="w", pady=(8,0))

        elif sys_name == "Darwin":
            tk.Label(frm,
                     text="Click below to install a macOS LaunchAgent.\n"
                          "The app will open automatically at login.",
                     font=FB, bg=BG, fg=SEC, justify="left").pack(anchor="w")
            primary_btn(frm, "⚙  Install LaunchAgent (macOS)",
                        self._setup_macos).pack(anchor="w", pady=(12,0))
            ghost_btn(frm, "Remove LaunchAgent",
                      self._remove_macos).pack(anchor="w", pady=(8,0))

        else:
            tk.Label(frm,
                     text="Add the following to your session startup\n"
                          "(GNOME: Startup Applications, KDE: Autostart):",
                     font=FB, bg=BG, fg=SEC, justify="left").pack(anchor="w")
            cmd = f"python3 {os.path.abspath(__file__)}"
            cmd_lbl = tk.Label(frm, text=cmd, font=FC, bg=SURFACE,
                               fg=ACCENT, padx=8, pady=6, relief="solid", bd=1)
            cmd_lbl.pack(anchor="w", pady=(10,0), fill="x")
            ghost_btn(frm, "Copy command",
                      lambda: (self.clipboard_clear(), self.clipboard_append(cmd))).pack(
                anchor="w", pady=(8,0))

    def _setup_windows(self):
        bat = _startup_bat_path()
        try:
            os.makedirs(os.path.dirname(bat), exist_ok=True)
            with open(bat, "w") as f:
                f.write(f'@echo off\nstart "" pythonw "{os.path.abspath(__file__)}"\n')
            messagebox.showinfo("Done",
                f"Startup launcher created.\nApp will open at each Windows login.\n\n{bat}",
                parent=self)
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)

    def _setup_macos(self):
        plist = _macos_plist_path()
        try:
            os.makedirs(os.path.dirname(plist), exist_ok=True)
            with open(plist, "w") as f:
                f.write(_macos_plist_content())
            subprocess.run(["launchctl", "load", plist], check=False)
            messagebox.showinfo("Done",
                f"LaunchAgent installed.\nApp will open at each login.\n\n{plist}",
                parent=self)
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)

    def _remove_macos(self):
        plist = _macos_plist_path()
        if os.path.exists(plist):
            subprocess.run(["launchctl", "unload", plist], check=False)
            os.remove(plist)
            messagebox.showinfo("Removed", "LaunchAgent removed.", parent=self)
        else:
            messagebox.showinfo("Not found", "No LaunchAgent found.", parent=self)

    # ── Save ──────────────────────────────────────────────────────────────────
    def _save(self):
        self.cfg["project_name"]       = self.proj_var.get().strip() or "Team"
        self.cfg["sensitivity_filter"] = self.filter_var.get()
        self.cfg["greetings"]          = list(self.greet_lb.get(0, "end"))

        for w in self._api_widgets:
            api = w["api"]
            api["enabled"] = w["enabled_var"].get()
            if "key_var"     in w: api["api_key"] = w["key_var"].get().strip()
            if "country_var" in w: api["country"] = w["country_var"].get().strip().upper()

        save_config(self.cfg)
        self.on_save()
        self.destroy()


# ─────────────────────────────────────────────────────────────────────────────
#  STARTUP HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _startup_bat_path():
    return os.path.join(
        os.environ.get("APPDATA",""),
        r"Microsoft\Windows\Start Menu\Programs\Startup",
        "holiday_messenger.bat"
    )

def _macos_plist_path():
    return os.path.expanduser("~/Library/LaunchAgents/com.holidaymessenger.app.plist")

def _macos_plist_content():
    return (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"\n'
        '  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n'
        '<plist version="1.0"><dict>\n'
        '  <key>Label</key><string>com.holidaymessenger.app</string>\n'
        '  <key>ProgramArguments</key><array>\n'
        f'    <string>{sys.executable}</string>\n'
        f'    <string>{os.path.abspath(__file__)}</string>\n'
        '  </array>\n'
        '  <key>RunAtLoad</key><true/>\n'
        '</dict></plist>'
    )

# ─────────────────────────────────────────────────────────────────────────────
#  MAIN APP  — 4-step wizard
# ─────────────────────────────────────────────────────────────────────────────
STEP_NAMES = ["Greeting", "Occasion", "Link", "Message"]

class HolidayMessengerApp:
    def __init__(self, root: tk.Tk):
        self.root    = root
        self.cfg     = load_config()
        self.now     = datetime.now()
        self.step    = 0

        # Per-wizard state
        self.chosen_greeting  = tk.StringVar()
        self.chosen_holiday   = None   # dict
        self.chosen_link      = None   # dict or None
        self.holidays         = []
        self.links            = []
        self._req_id          = 0      # invalidates stale async results
        self._copy_job        = None

        root.title("Holiday Morning Messenger")
        root.configure(bg=BG)
        root.resizable(False, False)
        self._build_shell()
        self._goto(0)
        self._center()

    # ── Shell (header, progress, content, footer — always present) ───────────
    def _build_shell(self):
        # ── Header
        hdr = tk.Frame(self.root, bg=ACCENT, padx=18, pady=11)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🌅  Holiday Morning Messenger",
                 font=("Helvetica Neue",14,"bold"),
                 bg=ACCENT, fg="#FFF").pack(side="left")
        tk.Button(hdr, text="⚙", command=self._open_settings,
                  font=("Helvetica Neue",13), bg=ACCD, fg="#FFF",
                  activebackground=ACCD, activeforeground="#FFF",
                  relief="flat", padx=8, pady=4, cursor="hand2", bd=0).pack(side="right")

        # ── Date bar
        db = tk.Frame(self.root, bg=SURFACE, padx=18, pady=7)
        db.pack(fill="x")
        tk.Label(db, text=f"📅  {self.now.strftime('%A, %d %B %Y')}",
                 font=FB, bg=SURFACE, fg=SEC).pack(side="left")

        # ── Progress indicator
        self.progress_frm = tk.Frame(self.root, bg=BG, padx=18, pady=10)
        self.progress_frm.pack(fill="x")
        hline(self.root)

        # ── Swappable content
        self.content = tk.Frame(self.root, bg=BG, width=500)
        self.content.pack(fill="both", expand=True)
        self.content.pack_propagate(True)
        hline(self.root)

        # ── Footer nav
        self.footer = tk.Frame(self.root, bg=BG, padx=18, pady=10)
        self.footer.pack(fill="x")

    def _update_progress(self):
        for w in self.progress_frm.winfo_children():
            w.destroy()
        for i, name in enumerate(STEP_NAMES):
            active = (i == self.step)
            done   = (i < self.step)
            col    = ACCENT if active else (GREEN if done else MUT)
            circle = "●" if active else ("✓" if done else "○")
            tk.Label(self.progress_frm, text=f"{circle} {name}",
                     font=("Helvetica Neue", 10, "bold" if active else "normal"),
                     bg=BG, fg=col).pack(side="left", padx=(0, 16))

    def _clear(self):
        for w in self.content.winfo_children():
            w.destroy()
        for w in self.footer.winfo_children():
            w.destroy()

    def _goto(self, n: int):
        self._req_id += 1           # invalidate pending async calls
        self.step = n
        self._update_progress()
        self._clear()
        [self._step0_greeting,
         self._step1_holiday,
         self._step2_links,
         self._step3_message][n]()

    def _back(self):    self._goto(self.step - 1)
    def _advance(self): self._goto(self.step + 1)

    # ────────────────────────────────────────────────────────────────────────
    #  STEP 0 — GREETING
    # ────────────────────────────────────────────────────────────────────────
    def _step0_greeting(self):
        frm = self._content_frame()
        tk.Label(frm, text="Choose your greeting",
                 font=FT, bg=BG, fg=PRI).pack(anchor="w", pady=(0,4))
        tk.Label(frm,
                 text="Greetings you've used recently are labelled so you can vary them.",
                 font=FS, bg=BG, fg=MUT, wraplength=440).pack(anchor="w", pady=(0,12))

        greetings = self.cfg.get("greetings", [])
        recent    = self.cfg.get("recent_greetings", [])
        proj      = self.cfg.get("project_name","Team")
        day       = self.now.strftime("%A")

        # Default to least-recently-used
        not_recent = [g for g in greetings if g not in recent] or greetings
        if not self.chosen_greeting.get() or self.chosen_greeting.get() not in greetings:
            self.chosen_greeting.set(not_recent[0])

        for g in greetings:
            display   = g.format(project=proj, day=day)
            is_recent = g in recent[-3:]
            row       = tk.Frame(frm, bg=BG)
            row.pack(fill="x", pady=3)
            tk.Radiobutton(row, text=display, variable=self.chosen_greeting,
                           value=g, font=FB, bg=BG, fg=PRI,
                           activebackground=BG, selectcolor=BG, cursor="hand2"
                           ).pack(side="left")
            if is_recent:
                tk.Label(row, text="• used recently", font=FS,
                         bg=BG, fg=MUT).pack(side="left", padx=(8,0))

        # Footer
        primary_btn(self.footer, "Next  →", self._on_greeting_next).pack(side="right")
        ghost_btn(self.footer, "⚙  Settings", self._open_settings).pack(side="left")

    def _on_greeting_next(self):
        g = self.chosen_greeting.get()
        if not g:
            messagebox.showwarning("Greeting", "Please choose a greeting.", parent=self.root)
            return
        # Track recency
        recent = self.cfg.setdefault("recent_greetings", [])
        if g in recent: recent.remove(g)
        recent.append(g)
        self.cfg["recent_greetings"] = recent[-12:]
        save_config(self.cfg)
        self._advance()

    # ────────────────────────────────────────────────────────────────────────
    #  STEP 1 — HOLIDAY / OCCASION
    # ────────────────────────────────────────────────────────────────────────
    def _step1_holiday(self):
        frm = self._content_frame()
        tk.Label(frm, text="Choose today's occasion",
                 font=FT, bg=BG, fg=PRI).pack(anchor="w", pady=(0,4))

        self._hol_status = tk.Label(frm,
                                    text="Fetching from configured API sources…",
                                    font=FS, bg=BG, fg=MUT, wraplength=440)
        self._hol_status.pack(anchor="w", pady=(0,10))

        # Scrollable container
        wrap  = tk.Frame(frm, bg=BG, relief="solid", bd=1)
        wrap.pack(fill="both", expand=True)
        self._hol_canvas = tk.Canvas(wrap, bg=BG, bd=0, highlightthickness=0, height=280)
        sb = tk.Scrollbar(wrap, orient="vertical", command=self._hol_canvas.yview)
        self._hol_list = tk.Frame(self._hol_canvas, bg=BG)
        self._hol_list.bind("<Configure>",
            lambda e: self._hol_canvas.configure(scrollregion=self._hol_canvas.bbox("all")))
        self._hol_canvas.create_window((0,0), window=self._hol_list, anchor="nw")
        self._hol_canvas.configure(yscrollcommand=sb.set)
        self._hol_canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._hol_canvas.bind_all("<MouseWheel>",
            lambda e: self._hol_canvas.yview_scroll(-(e.delta//120), "units"))

        self._holiday_var = tk.IntVar(value=-1)

        # Footer
        self._hol_next_btn = primary_btn(self.footer, "Next  →", self._on_holiday_next)
        self._hol_next_btn.pack(side="right")
        ghost_btn(self.footer, "←  Back", self._back).pack(side="right", padx=(0,8))
        ghost_btn(self.footer, "↺ Refresh", self._load_holidays).pack(side="left")

        self._load_holidays()

    def _load_holidays(self):
        # Clear list
        for w in self._hol_list.winfo_children():
            w.destroy()
        self._hol_status.config(text="Fetching from configured API sources…", fg=MUT)
        self._holiday_var.set(-1)
        req_id = self._req_id

        def worker():
            hols = fetch_all_holidays(self.cfg, self.now)
            if self._req_id == req_id:
                self.root.after(0, lambda: self._populate_holidays(hols))

        threading.Thread(target=worker, daemon=True).start()

    def _populate_holidays(self, holidays: list):
        self.holidays = holidays
        self._holiday_var = tk.IntVar(value=-1)

        if not holidays:
            self._hol_status.config(
                text="No occasions found for today. Check your API settings or try Refresh.",
                fg=RED)
            return

        # Restore previous selection if still present
        self._hol_status.config(
            text=f"{len(holidays)} occasion{'s' if len(holidays)!=1 else ''} found today.",
            fg=GREEN)

        prev_idx = -1
        if self.chosen_holiday:
            for i, h in enumerate(holidays):
                if h["name"].lower() == self.chosen_holiday["name"].lower():
                    prev_idx = i
                    break
        if prev_idx < 0 and holidays:
            prev_idx = 0
        self._holiday_var.set(prev_idx)
        if prev_idx >= 0:
            self.chosen_holiday = holidays[prev_idx]

        for i, h in enumerate(holidays):
            self._render_holiday_row(i, h)

    def _render_holiday_row(self, idx: int, h: dict):
        is_warn = (h["sensitivity"] == "warn")
        rbg     = WARN_BG if is_warn else BG

        row = tk.Frame(self._hol_list, bg=rbg, padx=8, pady=6)
        row.pack(fill="x")
        tk.Frame(self._hol_list, bg=BORDER, height=1).pack(fill="x")

        rb = tk.Radiobutton(row, variable=self._holiday_var, value=idx,
                            command=lambda i=idx: self._select_holiday(i),
                            bg=rbg, activebackground=rbg, selectcolor=rbg, cursor="hand2")
        rb.pack(side="left")

        body = tk.Frame(row, bg=rbg)
        body.pack(side="left", fill="x", expand=True)

        title_row = tk.Frame(body, bg=rbg)
        title_row.pack(anchor="w", fill="x")
        tk.Label(title_row, text=h["name"], font=FB,
                 bg=rbg, fg=WARN_FG if is_warn else PRI).pack(side="left")
        if is_warn:
            tk.Label(title_row, text="  ⚠️ handle with care",
                     font=FS, bg=rbg, fg=WARN_FG).pack(side="left", padx=(6,0))

        tk.Label(body, text=f"Source: {h['source']}", font=FS,
                 bg=rbg, fg=MUT).pack(anchor="w")

        # Make whole row clickable
        for w in [row, body, rb] + list(title_row.winfo_children()) + list(body.winfo_children()):
            w.bind("<Button-1>", lambda e, i=idx: (self._holiday_var.set(i), self._select_holiday(i)))

    def _select_holiday(self, idx: int):
        if 0 <= idx < len(self.holidays):
            self.chosen_holiday = self.holidays[idx]

    def _on_holiday_next(self):
        sel = self._holiday_var.get()
        if sel < 0 or not self.holidays:
            # Allow proceeding without a holiday (generic day message)
            self.chosen_holiday = None
        else:
            self.chosen_holiday = self.holidays[sel]
        self._advance()

    # ────────────────────────────────────────────────────────────────────────
    #  STEP 2 — LINKS
    # ────────────────────────────────────────────────────────────────────────
    def _step2_links(self):
        frm = self._content_frame()
        name = self.chosen_holiday["name"] if self.chosen_holiday else "today"
        tk.Label(frm, text="Choose a link to include",
                 font=FT, bg=BG, fg=PRI).pack(anchor="w", pady=(0,4))
        tk.Label(frm, text=f"Finding links relevant to: {name}",
                 font=FS, bg=BG, fg=MUT).pack(anchor="w", pady=(0,12))

        self._link_status = tk.Label(frm, text="🔍  Searching…",
                                     font=FS, bg=BG, fg=MUT)
        self._link_status.pack(anchor="w", pady=(0,6))

        self._link_frm = tk.Frame(frm, bg=BG)
        self._link_frm.pack(fill="both", expand=True)
        self._link_var = tk.IntVar(value=-1)

        # Footer
        primary_btn(self.footer, "Next  →", self._advance).pack(side="right")
        ghost_btn(self.footer, "←  Back", self._back).pack(side="right", padx=(0,8))

        req_id = self._req_id

        def worker():
            lks = find_links(self.chosen_holiday) if self.chosen_holiday else []
            if self._req_id == req_id:
                self.root.after(0, lambda: self._populate_links(lks))

        threading.Thread(target=worker, daemon=True).start()

    def _populate_links(self, links: list):
        self.links         = links
        self.chosen_link   = None
        self._link_var     = tk.IntVar(value=-1)
        self._link_status.config(
            text=f"{len(links)} link{'s' if len(links)!=1 else ''} found." if links
            else "No links found. You can continue without one.",
            fg=GREEN if links else MUT)

        ICONS = {"wikipedia":"📖", "source":"🔗", "official":"🌐",
                 "related":"💡", "default":"🔗"}

        for i, lk in enumerate(links):
            self._render_link_row(i, lk, ICONS)

        # "No link" row
        no_row = tk.Frame(self._link_frm, bg=BG, padx=8, pady=6)
        no_row.pack(fill="x")
        tk.Frame(self._link_frm, bg=BORDER, height=1).pack(fill="x")
        tk.Radiobutton(no_row, text="No link — message only",
                       variable=self._link_var, value=-1,
                       command=lambda: setattr(self, "chosen_link", None),
                       font=FB, bg=BG, fg=SEC,
                       activebackground=BG, selectcolor=BG,
                       cursor="hand2").pack(anchor="w")

    def _render_link_row(self, idx: int, lk: dict, icons: dict):
        row = tk.Frame(self._link_frm, bg=BG, padx=8, pady=6)
        row.pack(fill="x")
        tk.Frame(self._link_frm, bg=BORDER, height=1).pack(fill="x")

        icon = icons.get(lk["type"], icons["default"])
        rb   = tk.Radiobutton(row, variable=self._link_var, value=idx,
                              command=lambda i=idx: self._select_link(i),
                              bg=BG, activebackground=BG, selectcolor=BG, cursor="hand2")
        rb.pack(side="left", anchor="n", pady=(2,0))

        body = tk.Frame(row, bg=BG)
        body.pack(side="left", fill="x", expand=True)
        tk.Label(body, text=f"{icon}  {lk['label'][:72]}",
                 font=FB, bg=BG, fg=PRI, wraplength=400, justify="left").pack(anchor="w")

        url_short = lk["url"][:75] + ("…" if len(lk["url"])>75 else "")
        url_lbl   = tk.Label(body, text=url_short, font=FC,
                             bg=BG, fg=ACCENT, cursor="hand2")
        url_lbl.pack(anchor="w")
        url_lbl.bind("<Button-1>", lambda e, u=lk["url"]: webbrowser.open(u))

        for w in [row, body] + list(body.winfo_children()):
            w.bind("<Button-1>", lambda e, i=idx: (self._link_var.set(i), self._select_link(i)))

    def _select_link(self, idx: int):
        if 0 <= idx < len(self.links):
            self.chosen_link = self.links[idx]

    # ────────────────────────────────────────────────────────────────────────
    #  STEP 3 — FINAL MESSAGE
    # ────────────────────────────────────────────────────────────────────────
    def _step3_message(self):
        frm = self._content_frame()
        tk.Label(frm, text="Your message",
                 font=FT, bg=BG, fg=PRI).pack(anchor="w", pady=(0,10))

        self.msg_txt = tk.Text(frm, font=FM, width=52, height=8,
                               relief="solid", bd=1, wrap="word",
                               bg=SURFACE, fg=PRI, insertbackground=PRI,
                               padx=12, pady=10)
        self.msg_txt.pack(fill="x")
        self.msg_txt.insert("1.0", self._build_message())
        tk.Label(frm, text="Edit above if needed, then copy to Slack.",
                 font=FS, bg=BG, fg=MUT).pack(anchor="w", pady=(6,0))

        # Footer
        self.copy_btn = primary_btn(self.footer, "📋  Copy message", self._copy)
        self.copy_btn.pack(side="right")
        ghost_btn(self.footer, "←  Back", self._back).pack(side="right", padx=(0,8))
        ghost_btn(self.footer, "↺  Start over", lambda: self._goto(0)).pack(side="left")

    def _build_message(self) -> str:
        proj     = self.cfg.get("project_name","Team")
        day      = self.now.strftime("%A")
        greeting = self.chosen_greeting.get()
        greeting = greeting.format(project=proj, day=day)

        if self.chosen_holiday:
            body = f"{greeting}, happy {self.chosen_holiday['name']}!"
        else:
            body = f"{greeting}, happy {day}!"

        if self.chosen_link:
            body += f"\n\n{self.chosen_link['icon']}  {self.chosen_link['label']}\n{self.chosen_link['url']}"

        return body

    def _copy(self):
        msg = self.msg_txt.get("1.0","end").strip()
        self.root.clipboard_clear()
        self.root.clipboard_append(msg)
        self.root.update()
        self.copy_btn.config(text="✅  Copied!", bg=GREEN)
        if self._copy_job:
            self.root.after_cancel(self._copy_job)
        self._copy_job = self.root.after(
            2500, lambda: self.copy_btn.config(text="📋  Copy message", bg=ACCENT))

    # ────────────────────────────────────────────────────────────────────────
    #  HELPERS
    # ────────────────────────────────────────────────────────────────────────
    def _content_frame(self) -> tk.Frame:
        frm = tk.Frame(self.content, bg=BG, padx=20, pady=14)
        frm.pack(fill="both", expand=True)
        return frm

    def _open_settings(self):
        SettingsWindow(self.root, self.cfg, on_save=lambda: setattr(self, "cfg", load_config()))

    def _center(self):
        self.root.update_idletasks()
        w  = self.root.winfo_reqwidth()
        h  = self.root.winfo_reqheight()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//3}")


# ─────────────────────────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
def main():
    root = tk.Tk()
    root.tk_setPalette(background=BG)
    try:
        style = ttk.Style()
        style.theme_use("clam")
    except Exception:
        pass
    HolidayMessengerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
