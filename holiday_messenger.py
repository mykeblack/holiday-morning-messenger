#!/usr/bin/env python3
"""
Myke's Morning Message
A polished Tkinter wizard for creating a cheerful morning Slack/message post.

Flow:
  1. Choose greeting
  2. Choose today's occasion
  3. Choose optional link
  4. Choose one or more relevant emojis
  5. Review/copy message

API sources are defined in config.json and created on first run.
"""

import json
import html
import os
import platform
import random
import re
import subprocess
import sys
import threading
import webbrowser
from datetime import datetime
from urllib.parse import quote_plus
from urllib.request import Request, urlopen

import tkinter as tk
from tkinter import ttk, messagebox

# ─────────────────────────────────────────────────────────────────────────────
#  PATHS / CONFIG
# ─────────────────────────────────────────────────────────────────────────────
def app_dir() -> str:
    """Return the folder used for editable user files such as config.json.

    In normal Python this is the script folder. In a PyInstaller EXE this is
    the folder containing the EXE, so config.json is created next to the app
    instead of inside PyInstaller's temporary extraction folder.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resource_path(relative_path: str) -> str:
    """Return a bundled asset path for both normal Python and PyInstaller."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(APP_DIR, relative_path)


APP_DIR = app_dir()
CONFIG_PATH = os.path.join(APP_DIR, "config.json")
ASSETS_DIR = resource_path("assets")
APP_ICON_PNG = resource_path(os.path.join("assets", "app_icon.png"))
APP_LOGO_PNG = resource_path(os.path.join("assets", "app_logo.png"))
APP_ICON_ICO = resource_path(os.path.join("assets", "app_icon.ico"))

DEFAULT_CONFIG = {
    "project_name": "Team",
    "sensitivity_filter": True,
    "greetings": [
        "Good morning everyone",
        "Morning all",
        "Morning {project} people",
        "Hey team",
        "Rise and shine, team",
        "Happy {day}, everyone",
    ],
    "recent_greetings": [],
    "apis": [
        {
            "id": "wikipedia",
            "name": "Wikipedia – Observances & Holidays",
            "type": "wikipedia",
            "language": "en",
            "enabled": True,
            "notes": "No API key required. Pulls observances from Wikipedia date pages.",
        },
        {
            "id": "nager",
            "name": "Nager.Date – Public Holidays",
            "type": "nager",
            "country": "GB",
            "enabled": True,
            "notes": "No API key required. Change 'country' to your ISO country code.",
        },
        {
            "id": "checkiday",
            "name": "Checkiday – Fun & Novelty Days",
            "type": "checkiday",
            "api_key": "",
            "enabled": True,
            "public_fallback": True,
            "notes": "Best source for fun days like National Mr. Potato Head Day. Add an APILayer Checkiday API key for reliable results; without one, the app uses a lightweight public-page fallback where possible.",
        },
        {
            "id": "calendarific",
            "name": "Calendarific",
            "type": "calendarific",
            "api_key": "",
            "country": "GB",
            "enabled": False,
            "notes": "Free API key at https://calendarific.com – supports 230+ countries.",
        },
        {
            "id": "abstractapi",
            "name": "Abstract API – Holidays",
            "type": "abstractapi",
            "api_key": "",
            "country": "GB",
            "enabled": False,
            "notes": "Free API key at https://app.abstractapi.com/api/holidays",
        },
    ],
}

# ─────────────────────────────────────────────────────────────────────────────
#  SENSITIVITY FILTER
# ─────────────────────────────────────────────────────────────────────────────
_HARD_BLOCK = [
    "holocaust", "genocide", "hiroshima", "nagasaki", "nakba",
    "day of mourning", "kristallnacht", "srebrenica",
    "victims of aggression", "rwandan genocide", "armenian genocide",
    "day of memory and sorrow", "victims of fascism",
    "day of national tragedy", "black ribbon day", "remembrance of victims",
]

_SOFT_WARN = [
    "remembrance", "memorial", "fallen", "cancer", "hiv", "aids",
    "human trafficking", "slavery", "violence", "abuse", "missing",
    "poverty", "hunger", "homelessness", "exploitation", "epidemic",
    "pandemic", "drug abuse", "leprosy",
]

_NOT_AN_OBSERVANCE = re.compile(
    r"^(january|february|march|april|may|june|july|august|september|"
    r"october|november|december|\d{4}|saint |st\.|feast of)",
    re.I,
)

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

# ─────────────────────────────────────────────────────────────────────────────
#  PRETTY THEME
# ─────────────────────────────────────────────────────────────────────────────
APP_TITLE = "Myke's morning message"
BG = "#F7F7FC"
CARD = "#FFFFFF"
CARD_SOFT = "#FBFAFF"
SURFACE = "#F1F0FF"
BORDER = "#E6E4F2"
BORDER_DARK = "#D6D2E8"
ACCENT = "#5B3DF5"
ACCENT_DARK = "#4327D7"
ACCENT_SOFT = "#EFEAFF"
GREEN = "#16A34A"
GREEN_SOFT = "#EAF8EF"
RED = "#DC2626"
WARN_BG = "#FFFBEB"
WARN_FG = "#92400E"
PRI = "#17142A"
SEC = "#625F78"
MUT = "#9994AE"
LINK = "#4F46E5"

FONT_TITLE = ("Segoe UI", 18, "bold")
FONT_H1 = ("Segoe UI", 16, "bold")
FONT_H2 = ("Segoe UI", 12, "bold")
FONT_BODY = ("Segoe UI", 11)
FONT_SMALL = ("Segoe UI", 9)
FONT_MONO = ("Consolas", 10)
FONT_EMOJI = ("Segoe UI Emoji", 18)
FONT_EMOJI_SMALL = ("Segoe UI Emoji", 12)

STEP_NAMES = ["Greeting", "Occasion", "Link", "Emoji", "Message"]

# ─────────────────────────────────────────────────────────────────────────────
#  SMALL UI HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def rounded_label(parent, text, bg, fg, font=FONT_SMALL, padx=8, pady=3):
    # Tk doesn't have true rounded labels; this gives a clean badge-like look.
    return tk.Label(parent, text=text, bg=bg, fg=fg, font=font, padx=padx, pady=pady)


def _button(parent, text, command, kind="primary", font=FONT_BODY, padx=16, pady=8, **kw):
    if kind == "primary":
        bg, fg, abg, afg = ACCENT, "#FFFFFF", ACCENT_DARK, "#FFFFFF"
    elif kind == "danger":
        bg, fg, abg, afg = "#FEF2F2", RED, "#FEE2E2", RED
    elif kind == "success":
        bg, fg, abg, afg = GREEN, "#FFFFFF", "#15803D", "#FFFFFF"
    else:
        bg, fg, abg, afg = "#FFFFFF", SEC, ACCENT_SOFT, PRI

    return tk.Button(
        parent,
        text=text,
        command=command,
        font=font,
        bg=bg,
        fg=fg,
        activebackground=abg,
        activeforeground=afg,
        relief="flat",
        padx=padx,
        pady=pady,
        cursor="hand2",
        bd=0,
        highlightthickness=1,
        highlightbackground=BORDER,
        **kw,
    )


def primary_btn(parent, text, cmd, **kw):
    return _button(parent, text, cmd, "primary", **kw)


def ghost_btn(parent, text, cmd, **kw):
    return _button(parent, text, cmd, "ghost", **kw)


def danger_btn(parent, text, cmd, **kw):
    return _button(parent, text, cmd, "danger", **kw)


def hline(parent, color=BORDER, **kw):
    tk.Frame(parent, bg=color, height=1).pack(fill="x", **kw)


def make_card(parent, padx=18, pady=14):
    card = tk.Frame(parent, bg=CARD, padx=padx, pady=pady, highlightthickness=1, highlightbackground=BORDER)
    return card

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG / HTTP HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def load_config() -> dict:
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                cfg = json.load(f)
            for k, v in DEFAULT_CONFIG.items():
                cfg.setdefault(k, v)

            # Upgrade older config files by adding any newly introduced API source.
            existing_api_ids = {api.get("id") for api in cfg.get("apis", [])}
            for api in DEFAULT_CONFIG.get("apis", []):
                if api.get("id") not in existing_api_ids:
                    cfg.setdefault("apis", []).append(json.loads(json.dumps(api)))
            return cfg
        except Exception:
            pass
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    save_config(cfg)
    return cfg


def save_config(cfg: dict):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


_UA = {"User-Agent": "MykesMorningMessage/3.0 (https://github.com/mykeblack/holiday-morning-messenger)"}


def fetch_json(url: str, timeout: int = 9, headers: dict | None = None):
    try:
        all_headers = dict(_UA)
        if headers:
            all_headers.update(headers)
        req = Request(url, headers=all_headers)
        with urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception:
        return None


def fetch_text(url: str, timeout: int = 9, headers: dict | None = None):
    try:
        all_headers = dict(_UA)
        if headers:
            all_headers.update(headers)
        req = Request(url, headers=all_headers)
        with urlopen(req, timeout=timeout) as resp:
            return resp.read().decode("utf-8", errors="replace")
    except Exception:
        return None


def sensitivity(name: str) -> str:
    lo = name.lower()
    if any(k in lo for k in _HARD_BLOCK):
        return "block"
    if any(k in lo for k in _SOFT_WARN):
        return "warn"
    return "ok"


def _holiday(name, source, url="", sens=None, fun_score: int | None = None) -> dict:
    if sens is None:
        sens = sensitivity(name)
    item = {"name": name, "source": source, "url": url, "sensitivity": sens}
    if fun_score is not None:
        item["fun_score"] = fun_score
    return item


FUN_KEYWORDS = [
    "potato", "bubble tea", "cookie", "cake", "chocolate", "pizza", "donut", "coffee", "tea",
    "toy", "game", "jazz", "music", "dance", "cat", "dog", "pet", "animal", "smile",
    "kindness", "friendship", "star wars", "superhero", "comic", "ice cream", "waffle",
    "pancake", "taco", "burger", "cheese", "cupcake", "silly", "fun", "party", "laugh",
]
SERIOUS_KEYWORDS = [
    "memorial", "remembrance", "victims", "disease", "cancer", "aids", "war", "mourning",
    "violence", "abuse", "poverty", "hunger", "trafficking", "genocide", "epidemic", "pandemic",
]


def fun_priority(holiday: dict) -> int:
    """Higher means more likely to be the light-hearted item people want first."""
    name = holiday.get("name", "").lower()
    source = holiday.get("source", "").lower()
    score = int(holiday.get("fun_score", 0) or 0)

    if "checkiday" in source:
        score += 70
    if "fun" in source or "novelty" in source:
        score += 45
    if name.startswith("national "):
        score += 18
    if name.endswith(" day") or " day" in name:
        score += 10
    score += sum(18 for k in FUN_KEYWORDS if k in name)
    score -= sum(28 for k in SERIOUS_KEYWORDS if k in name)
    if holiday.get("sensitivity") == "warn":
        score -= 100
    if "nager" in source:
        score -= 10
    return score

# ─────────────────────────────────────────────────────────────────────────────
#  API FETCHERS
# ─────────────────────────────────────────────────────────────────────────────
def fetch_wikipedia(date: datetime, language: str = "en") -> list[dict]:
    page = f"{MONTH_NAMES[date.month - 1]}_{date.day}"
    base = f"https://{language}.wikipedia.org/w/api.php"

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

    content = fetch_json(f"{base}?action=parse&page={page}&prop=wikitext&section={section_idx}&format=json")
    if not content:
        return []
    wikitext = ((content.get("parse") or {}).get("wikitext") or {}).get("*", "")

    results, seen = [], set()
    for line in wikitext.split("\n"):
        line = line.strip()
        if not line.startswith("*"):
            continue
        m = re.search(r"\[\[([^\]|#]+?)(?:\|[^\]]+)?\]\]", line)
        if not m:
            continue
        name = m.group(1).strip()
        if name in seen or re.match(r"^\d{4}$", name) or _NOT_AN_OBSERVANCE.match(name) or len(name) < 5:
            continue
        sens = sensitivity(name)
        if sens == "block":
            continue
        seen.add(name)
        wiki_url = f"https://{language}.wikipedia.org/wiki/{name.replace(' ', '_')}"
        results.append(_holiday(name, "Wikipedia", wiki_url, sens))
    return results


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


def fetch_calendarific(date: datetime, api_key: str, country: str = "GB") -> list[dict]:
    if not api_key.strip():
        return []
    url = (
        "https://calendarific.com/api/v2/holidays"
        f"?api_key={api_key}&country={country}&year={date.year}&month={date.month}&day={date.day}"
    )
    data = fetch_json(url)
    if not data:
        return []
    hols = ((data.get("response") or {}).get("holidays")) or []
    results = []
    for h in hols:
        name = h.get("name", "")
        if name:
            sens = sensitivity(name)
            if sens != "block":
                results.append(_holiday(name, "Calendarific", "", sens))
    return results


def fetch_abstractapi(date: datetime, api_key: str, country: str = "GB") -> list[dict]:
    if not api_key.strip():
        return []
    url = (
        "https://holidays.abstractapi.com/v1/"
        f"?api_key={api_key}&country={country}&year={date.year}&month={date.month}&day={date.day}"
    )
    data = fetch_json(url)
    if not isinstance(data, list):
        return []
    results = []
    for h in data:
        name = h.get("name", "")
        if name:
            sens = sensitivity(name)
            if sens != "block":
                results.append(_holiday(name, "Abstract API", "", sens))
    return results


def _extract_checkiday_events(data) -> list[dict]:
    """Accept a few response shapes because marketplace examples can differ by plan/version."""
    if not data:
        return []
    if isinstance(data, list):
        events = data
    elif isinstance(data, dict):
        events = (
            data.get("events")
            or data.get("holidays")
            or data.get("data")
            or ((data.get("response") or {}).get("events"))
            or ((data.get("response") or {}).get("holidays"))
            or []
        )
    else:
        events = []

    results = []
    for event in events:
        if not isinstance(event, dict):
            continue
        name = (event.get("name") or event.get("title") or event.get("event") or "").strip()
        if not name:
            continue
        url = event.get("url") or event.get("link") or ""
        event_id = event.get("id") or event.get("uid")
        if not url and event_id:
            url = f"https://www.checkiday.com/{event_id}"
        sens = sensitivity(name)
        if sens != "block":
            results.append(_holiday(name, "Checkiday", url, sens, fun_score=90))
    return results


def fetch_checkiday(date: datetime, api_key: str = "", public_fallback: bool = True) -> list[dict]:
    """Fetch fun/novelty days from Checkiday.

    The APILayer Checkiday API requires an `apikey` header. The public fallback is best-effort
    and only intended to keep the app useful without a key; the API is the reliable route.
    """
    results = []
    date_str = date.strftime("%Y-%m-%d")

    if api_key.strip():
        # APILayer documentation lists /events with a date query parameter.
        api_urls = [
            f"https://api.apilayer.com/checkiday/events?date={date_str}&adult=false",
            f"https://api.apilayer.com/checkiday/events?adult=false&date={date_str}",
        ]
        for url in api_urls:
            data = fetch_json(url, headers={"apikey": api_key.strip()})
            results = _extract_checkiday_events(data)
            if results:
                return results

    if public_fallback:
        results = fetch_checkiday_public(date)
        if results:
            return results

    return fetch_curated_fun_days(date)


def fetch_checkiday_public(date: datetime) -> list[dict]:
    # Checkiday's public homepage shows today's holidays. Use it only for today's date.
    today = datetime.now()
    if date.date() != today.date():
        return []
    page = fetch_text("https://www.checkiday.com/")
    if not page:
        return []

    results, seen = [], set()
    # The public page commonly renders event titles as headings. Keep parsing conservative.
    for raw in re.findall(r"<h2[^>]*>(.*?)</h2>", page, flags=re.I | re.S):
        name = re.sub(r"<[^>]+>", "", raw)
        name = html.unescape(name).strip()
        name = re.sub(r"\s+", " ", name)
        if not name or name.lower() in seen or len(name) < 4:
            continue
        sens = sensitivity(name)
        if sens == "block":
            continue
        seen.add(name.lower())
        results.append(_holiday(name, "Checkiday public", "https://www.checkiday.com/", sens, fun_score=75))
    return results


CURATED_FUN_DAYS = {
    "04-30": [
        ("National Mr. Potato Head Day", "https://www.checkiday.com/44c09fd393ca7b3e36ac35db863f3925/national-mr-potato-head-day"),
        ("National Bubble Tea Day", "https://www.checkiday.com/"),
        ("National Oatmeal Cookie Day", "https://www.checkiday.com/"),
        ("National Raisin Day", "https://www.checkiday.com/"),
    ],
}


def fetch_curated_fun_days(date: datetime) -> list[dict]:
    """Tiny built-in safety net for favourite novelty days if the web/API is unavailable."""
    key = date.strftime("%m-%d")
    return [_holiday(name, "Built-in fun days", url, fun_score=65) for name, url in CURATED_FUN_DAYS.get(key, [])]


def fetch_all_holidays(cfg: dict, date: datetime) -> list[dict]:
    results = []
    for api in cfg.get("apis", []):
        if not api.get("enabled"):
            continue
        try:
            t = api["type"]
            if t == "wikipedia":
                results += fetch_wikipedia(date, api.get("language", "en"))
            elif t == "nager":
                results += fetch_nager(date, api.get("country", "GB"))
            elif t == "checkiday":
                results += fetch_checkiday(date, api.get("api_key", ""), api.get("public_fallback", True))
            elif t == "calendarific":
                results += fetch_calendarific(date, api.get("api_key", ""), api.get("country", "GB"))
            elif t == "abstractapi":
                results += fetch_abstractapi(date, api.get("api_key", ""), api.get("country", "GB"))
        except Exception:
            pass

    seen, unique = set(), []
    for h in results:
        key = h["name"].lower()
        if key not in seen:
            seen.add(key)
            unique.append(h)
    unique.sort(key=lambda h: (0 if h["sensitivity"] == "ok" else 1, -fun_priority(h), h["name"]))
    return unique

# ─────────────────────────────────────────────────────────────────────────────
#  LINK DISCOVERY
# ─────────────────────────────────────────────────────────────────────────────
def find_links(holiday: dict) -> list[dict]:
    name = holiday["name"]
    links, seen = [], set()

    def add(label, url, icon, ltype):
        if url and url not in seen:
            seen.add(url)
            links.append({"label": label, "url": url, "icon": icon, "type": ltype})

    q = quote_plus(name)
    data = fetch_json(
        f"https://en.wikipedia.org/w/api.php?action=opensearch&search={q}&limit=3&namespace=0&format=json"
    )
    if data and len(data) >= 4:
        for title, url in zip(data[1], data[3]):
            add(f"Wikipedia: {title}", url, "📖", "wikipedia")

    if holiday.get("url"):
        add(f"{holiday['source']} page", holiday["url"], "🔗", "source")

    ddg = fetch_json(f"https://api.duckduckgo.com/?q={q}&format=json&no_redirect=1&no_html=1&skip_disambig=1")
    if ddg:
        if ddg.get("AbstractURL"):
            src = ddg.get("AbstractSource") or "Info"
            add(src, ddg["AbstractURL"], "🌐", "official")
        if ddg.get("OfficialSite"):
            add("Official website", ddg["OfficialSite"], "🌐", "official")
        for topic in (ddg.get("RelatedTopics") or [])[:4]:
            if isinstance(topic, dict) and topic.get("FirstURL"):
                text = re.sub(r"<[^>]+>", "", topic.get("Text") or "")[:72]
                add(text or "Related link", topic["FirstURL"], "💡", "related")

    slug = name.lower().replace("'", "").replace(",", "")
    slug = re.sub(r"[^a-z0-9]+", "-", slug).strip("-")
    if any(w in name.lower() for w in ["world ", "international ", "global "]):
        add("UN Observance page", f"https://www.un.org/en/observances/{slug}", "🇺🇳", "official")
        if any(w in name.lower() for w in ["health", "disease", "cancer", "aids", "diabetes", "blood", "nurse", "tobacco", "patient"]):
            add("WHO Campaign page", f"https://www.who.int/campaigns/{slug}", "🏥", "official")

    return links[:7]

# ─────────────────────────────────────────────────────────────────────────────
#  EMOJI SUGGESTIONS
# ─────────────────────────────────────────────────────────────────────────────
BASE_EMOJIS = ["✨", "🌞", "😊", "🎉", "👏", "💜", "🌍", "💫", "🙌", "⭐"]
KEYWORD_EMOJIS = [
    (["jazz", "music", "song", "dance", "piano", "guitar", "violin", "sax"], ["🎷", "🎶", "🎹", "🎺", "🎸", "🥁", "🎵", "💃"]),
    (["children", "child", "kid", "youth"], ["🧒", "👧", "👦", "🎈", "🧸", "🍭", "🌈", "💛"]),
    (["honesty", "truth", "integrity"], ["🤝", "💬", "✅", "🫶", "💙", "✨", "📣", "🌟"]),
    (["earth", "environment", "climate", "nature", "tree", "wildlife", "animal", "veterinary", "bird"], ["🌍", "🌱", "🌳", "🐾", "🦋", "🌿", "🐶", "🐱"]),
    (["book", "read", "literacy", "poetry", "author", "story"], ["📚", "📖", "✍️", "📝", "💡", "🎓", "🖋️", "✨"]),
    (["science", "space", "math", "technology", "engineering"], ["🔬", "🚀", "🧪", "💻", "🛰️", "🧠", "⚙️", "✨"]),
    (["potato", "spud", "tater"], ["🥔", "🧸", "📺", "🎩", "👓", "🥸", "🎉", "✨"]),
    (["bubble tea", "boba"], ["🧋", "😋", "🎉", "✨", "💜"]),
    (["cookie", "oatmeal", "raisin"], ["🍪", "😋", "🥛", "🎉", "✨"]),
    (["food", "coffee", "tea", "cake", "chocolate", "pizza"], ["☕", "🍰", "🍫", "🍕", "🥐", "🍽️", "😋", "🎉"]),
    (["health", "nurse", "doctor", "blood", "heart", "mental"], ["💚", "❤️", "🩺", "🏥", "💪", "🌱", "🤗", "✨"]),
    (["peace", "friendship", "kindness", "smile", "family"], ["🕊️", "🤝", "😊", "💞", "🫶", "🌈", "✨", "🙌"]),
    (["star", "wars", "space"], ["⭐", "🌌", "🚀", "🛸", "✨", "🌙", "👾", "💫"]),
    (["cat"], ["🐱", "😺", "🐾", "💛", "✨"]),
    (["dog"], ["🐶", "🐕", "🐾", "💛", "✨"]),
    (["veterinary", "vet", "animal"], ["🐾", "🐶", "🐱", "🐰", "🩺", "💚", "✨"]),
]


def emoji_suggestions_for(holiday_name: str) -> list[str]:
    name = (holiday_name or "").lower()
    selected = []
    for keywords, emojis in KEYWORD_EMOJIS:
        if any(k in name for k in keywords):
            selected.extend(emojis)
    selected.extend(BASE_EMOJIS)
    # unique, preserve order
    out = []
    for e in selected:
        if e not in out:
            out.append(e)
    return out[:24]

# ─────────────────────────────────────────────────────────────────────────────
#  SETTINGS WINDOW
# ─────────────────────────────────────────────────────────────────────────────
class SettingsWindow(tk.Toplevel):
    def __init__(self, parent, cfg: dict, on_save):
        super().__init__(parent)
        self.title("Settings")
        apply_window_icon(self)
        self.cfg = cfg
        self.on_save = on_save
        self.configure(bg=BG)
        self.resizable(False, False)
        self.grab_set()
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=ACCENT, padx=18, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Settings", bg=ACCENT, fg="#FFFFFF", font=FONT_H1).pack(anchor="w")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=12)
        self._tab_general(nb)
        self._tab_apis(nb)
        self._tab_greetings(nb)
        self._tab_startup(nb)

        footer = tk.Frame(self, bg=BG, padx=14, pady=12)
        footer.pack(fill="x")
        primary_btn(footer, "Save & close", self._save).pack(side="right", padx=(8, 0))
        ghost_btn(footer, "Cancel", self.destroy).pack(side="right")

    def _tab_general(self, nb):
        frm = tk.Frame(nb, bg=CARD, padx=18, pady=18)
        nb.add(frm, text="General")
        tk.Label(frm, text="Project / team name", font=FONT_H2, bg=CARD, fg=PRI).pack(anchor="w")
        self.proj_var = tk.StringVar(value=self.cfg.get("project_name", "Team"))
        tk.Entry(frm, textvariable=self.proj_var, font=FONT_BODY, width=34, relief="solid", bd=1).pack(anchor="w", pady=(5, 16))
        self.filter_var = tk.BooleanVar(value=self.cfg.get("sensitivity_filter", True))
        tk.Checkbutton(
            frm,
            text="Filter out culturally insensitive days",
            variable=self.filter_var,
            font=FONT_BODY,
            bg=CARD,
            fg=PRI,
            activebackground=CARD,
            selectcolor=CARD,
        ).pack(anchor="w")

    def _tab_apis(self, nb):
        outer = tk.Frame(nb, bg=CARD)
        nb.add(outer, text="API Sources")
        canvas = tk.Canvas(outer, bg=CARD, bd=0, highlightthickness=0, height=360, width=520)
        sb = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        frm = tk.Frame(canvas, bg=CARD, padx=18, pady=12)
        frm.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frm, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self._api_widgets = []
        for api in self.cfg.get("apis", []):
            self._api_card(frm, api)

    def _api_card(self, parent, api):
        card = make_card(parent, padx=12, pady=10)
        card.configure(bg=CARD_SOFT)
        card.pack(fill="x", pady=6)
        ev = tk.BooleanVar(value=api.get("enabled", False))
        tk.Checkbutton(
            card,
            text=api["name"],
            variable=ev,
            font=FONT_H2,
            bg=CARD_SOFT,
            fg=PRI,
            activebackground=CARD_SOFT,
            selectcolor=CARD_SOFT,
        ).pack(anchor="w")
        if api.get("notes"):
            tk.Label(card, text=api["notes"], font=FONT_SMALL, bg=CARD_SOFT, fg=MUT, wraplength=440, justify="left").pack(anchor="w", pady=(3, 4))
        widgets = {"enabled_var": ev, "api": api}
        if api["type"] in ("calendarific", "abstractapi", "checkiday"):
            kv = tk.StringVar(value=api.get("api_key", ""))
            row = tk.Frame(card, bg=CARD_SOFT)
            row.pack(fill="x", pady=3)
            tk.Label(row, text="API key:", font=FONT_SMALL, bg=CARD_SOFT, fg=SEC).pack(side="left")
            tk.Entry(row, textvariable=kv, font=FONT_MONO, width=38, relief="solid", bd=1, show="*").pack(side="left", padx=(6, 0))
            widgets["key_var"] = kv
        if api["type"] in ("nager", "calendarific", "abstractapi"):
            cv = tk.StringVar(value=api.get("country", "GB"))
            row = tk.Frame(card, bg=CARD_SOFT)
            row.pack(fill="x", pady=3)
            tk.Label(row, text="Country ISO code:", font=FONT_SMALL, bg=CARD_SOFT, fg=SEC).pack(side="left")
            tk.Entry(row, textvariable=cv, font=FONT_BODY, width=6, relief="solid", bd=1).pack(side="left", padx=(6, 0))
            widgets["country_var"] = cv
        self._api_widgets.append(widgets)

    def _tab_greetings(self, nb):
        frm = tk.Frame(nb, bg=CARD, padx=18, pady=18)
        nb.add(frm, text="Greetings")
        tk.Label(frm, text="Greeting templates", font=FONT_H1, bg=CARD, fg=PRI).pack(anchor="w")
        tk.Label(frm, text="Use {project} and {day} as placeholders.", font=FONT_SMALL, bg=CARD, fg=MUT).pack(anchor="w", pady=(2, 10))
        list_frm = tk.Frame(frm, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        list_frm.pack(fill="x")
        sb = tk.Scrollbar(list_frm)
        sb.pack(side="right", fill="y")
        self.greet_lb = tk.Listbox(
            list_frm,
            font=FONT_BODY,
            yscrollcommand=sb.set,
            bg=CARD,
            fg=PRI,
            selectbackground=ACCENT,
            selectforeground="#FFF",
            relief="flat",
            height=8,
            activestyle="none",
        )
        self.greet_lb.pack(side="left", fill="both", expand=True)
        sb.config(command=self.greet_lb.yview)
        for g in self.cfg.get("greetings", []):
            self.greet_lb.insert("end", g)
        btns = tk.Frame(frm, bg=CARD)
        btns.pack(fill="x", pady=(10, 0))
        ghost_btn(btns, "＋ Add", self._add_greeting, font=FONT_SMALL, padx=10, pady=5).pack(side="left", padx=(0, 5))
        ghost_btn(btns, "✎ Edit", self._edit_greeting, font=FONT_SMALL, padx=10, pady=5).pack(side="left", padx=(0, 5))
        danger_btn(btns, "✕ Delete", self._del_greeting, font=FONT_SMALL, padx=10, pady=5).pack(side="left")

    def _add_greeting(self):
        self._greeting_dialog("", True)

    def _edit_greeting(self):
        sel = self.greet_lb.curselection()
        if not sel:
            messagebox.showinfo("Edit", "Select a greeting first.", parent=self)
            return
        self._greeting_dialog(self.greet_lb.get(sel[0]), False, sel[0])

    def _del_greeting(self):
        sel = self.greet_lb.curselection()
        if not sel:
            return
        if self.greet_lb.size() <= 1:
            messagebox.showwarning("Delete", "Keep at least one greeting.", parent=self)
            return
        self.greet_lb.delete(sel[0])

    def _greeting_dialog(self, text, is_new, idx=None):
        win = tk.Toplevel(self)
        win.title("Add greeting" if is_new else "Edit greeting")
        win.configure(bg=BG)
        win.resizable(False, False)
        win.grab_set()
        card = make_card(win)
        card.pack(fill="both", expand=True, padx=14, pady=14)
        tk.Label(card, text="Greeting text", font=FONT_H2, bg=CARD, fg=PRI).pack(anchor="w")
        var = tk.StringVar(value=text)
        tk.Entry(card, textvariable=var, font=FONT_BODY, width=44, relief="solid", bd=1).pack(pady=(5, 6))
        tk.Label(card, text="Placeholders: {project}  {day}", font=FONT_SMALL, bg=CARD, fg=MUT).pack(anchor="w")
        row = tk.Frame(card, bg=CARD)
        row.pack(fill="x", pady=(14, 0))

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

        primary_btn(row, "Save", save).pack(side="right", padx=(8, 0))
        ghost_btn(row, "Cancel", win.destroy).pack(side="right")

    def _tab_startup(self, nb):
        frm = tk.Frame(nb, bg=CARD, padx=18, pady=18)
        nb.add(frm, text="Startup")
        sys_name = platform.system()
        tk.Label(frm, text=f"Detected OS: {sys_name}", font=FONT_H2, bg=CARD, fg=PRI).pack(anchor="w", pady=(0, 12))
        if sys_name == "Windows":
            tk.Label(frm, text="Create a launcher in your Windows Startup folder.", font=FONT_BODY, bg=CARD, fg=SEC).pack(anchor="w")
            primary_btn(frm, "Install startup launcher", self._setup_windows).pack(anchor="w", pady=(12, 0))
        else:
            cmd = f"{sys.executable} {os.path.abspath(__file__)}"
            tk.Label(frm, text="Add this command to your session startup:", font=FONT_BODY, bg=CARD, fg=SEC).pack(anchor="w")
            tk.Label(frm, text=cmd, font=FONT_MONO, bg=ACCENT_SOFT, fg=ACCENT, padx=8, pady=6).pack(anchor="w", fill="x", pady=(8, 0))

    def _setup_windows(self):
        bat = _startup_bat_path()
        try:
            os.makedirs(os.path.dirname(bat), exist_ok=True)
            with open(bat, "w", encoding="utf-8") as f:
                f.write(f'@echo off\nstart "" "{sys.executable.replace("python.exe", "pythonw.exe")}" "{os.path.abspath(__file__)}"\n')
            messagebox.showinfo("Done", f"Startup launcher created:\n\n{bat}", parent=self)
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)

    def _save(self):
        self.cfg["project_name"] = self.proj_var.get().strip() or "Team"
        self.cfg["sensitivity_filter"] = self.filter_var.get()
        self.cfg["greetings"] = list(self.greet_lb.get(0, "end"))
        for w in self._api_widgets:
            api = w["api"]
            api["enabled"] = w["enabled_var"].get()
            if "key_var" in w:
                api["api_key"] = w["key_var"].get().strip()
            if "country_var" in w:
                api["country"] = w["country_var"].get().strip().upper()
        save_config(self.cfg)
        self.on_save()
        self.destroy()


def _startup_bat_path():
    return os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\Start Menu\Programs\Startup", "mykes_morning_message.bat")

# ─────────────────────────────────────────────────────────────────────────────
#  WINDOW ICON / LOGO HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def apply_window_icon(win):
    """Apply the custom app icon to a Tk/Toplevel window when available.

    PyInstaller uses the .ico file for the EXE icon. Tkinter is more reliable
    with .png at runtime, especially in --onefile builds, so the window and
    taskbar icon are set with PhotoImage.
    """
    try:
        if os.path.exists(APP_ICON_PNG):
            icon = tk.PhotoImage(file=APP_ICON_PNG)
            win.iconphoto(True, icon)
            win._app_icon_ref = icon  # keep a reference so Tk does not discard it
            return
    except Exception:
        pass

    try:
        if platform.system() == "Windows" and os.path.exists(APP_ICON_ICO):
            win.iconbitmap(APP_ICON_ICO)
    except Exception:
        pass

# ─────────────────────────────────────────────────────────────────────────────
#  MAIN APP
# ─────────────────────────────────────────────────────────────────────────────
class MorningMessageApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.cfg = load_config()
        self.now = datetime.now()
        self.step = 0
        self.chosen_greeting = tk.StringVar()
        self.chosen_holiday = None
        self.chosen_link = None
        self.holidays = []
        self.links = []
        self.selected_emojis = []
        self.emoji_vars = {}
        self._req_id = 0
        self._copy_job = None

        root.title(APP_TITLE)
        apply_window_icon(root)
        root.configure(bg=BG)
        root.resizable(True, True)
        root.minsize(720, 760)
        self._setup_styles()
        self._build_shell()
        self._goto(0)
        self._center()

    def _setup_styles(self):
        try:
            style = ttk.Style()
            style.theme_use("clam")
            style.configure("TNotebook", background=BG, borderwidth=0)
            style.configure("TNotebook.Tab", padding=(12, 7), font=FONT_SMALL)
        except Exception:
            pass

    def _build_shell(self):
        hdr = tk.Frame(self.root, bg=ACCENT, padx=24, pady=16)
        hdr.pack(fill="x")
        self._header_logo = None
        if os.path.exists(APP_LOGO_PNG):
            try:
                self._header_logo = tk.PhotoImage(file=APP_LOGO_PNG)
                tk.Label(hdr, image=self._header_logo, bg=ACCENT).pack(side="left", padx=(0, 12))
            except Exception:
                tk.Label(hdr, text="🌅", bg=ACCENT, fg="#FFFFFF", font=("Segoe UI Emoji", 20, "bold")).pack(side="left", padx=(0, 10))
        else:
            tk.Label(hdr, text="🌅", bg=ACCENT, fg="#FFFFFF", font=("Segoe UI Emoji", 20, "bold")).pack(side="left", padx=(0, 10))
        tk.Label(hdr, text=APP_TITLE, font=FONT_TITLE, bg=ACCENT, fg="#FFFFFF").pack(side="left")
        tk.Button(
            hdr,
            text="⚙",
            command=self._open_settings,
            font=("Segoe UI Symbol", 15),
            bg=ACCENT_DARK,
            fg="#FFFFFF",
            activebackground=ACCENT_DARK,
            activeforeground="#FFFFFF",
            relief="flat",
            padx=12,
            pady=7,
            cursor="hand2",
            bd=0,
        ).pack(side="right")

        date_bar = tk.Frame(self.root, bg=CARD, padx=24, pady=10)
        date_bar.pack(fill="x")
        tk.Label(date_bar, text=f"📅  {self.now.strftime('%A, %d %B %Y')}", font=FONT_BODY, bg=CARD, fg=SEC).pack(side="left")

        self.progress_frm = tk.Frame(self.root, bg=BG, padx=24, pady=14)
        self.progress_frm.pack(fill="x")

        self.content = tk.Frame(self.root, bg=BG, width=680, height=500)
        self.content.pack(fill="both", expand=True, padx=20)
        self.content.pack_propagate(False)

        self.footer = tk.Frame(self.root, bg=BG, padx=24, pady=16)
        self.footer.pack(fill="x")

    def _center(self):
        self.root.update_idletasks()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w, h = 760, 800
        self.root.geometry(f"{w}x{h}+{(sw - w) // 2}+{max(20, (sh - h) // 4)}")

    def _update_progress(self):
        for w in self.progress_frm.winfo_children():
            w.destroy()
        for i, name in enumerate(STEP_NAMES):
            done = i < self.step
            active = i == self.step
            if done:
                bg, fg, symbol = GREEN_SOFT, GREEN, "✓"
            elif active:
                bg, fg, symbol = ACCENT, "#FFFFFF", str(i + 1)
            else:
                bg, fg, symbol = "#ECEAF3", MUT, str(i + 1)
            item = tk.Frame(self.progress_frm, bg=BG)
            item.pack(side="left", padx=(0, 16))
            tk.Label(item, text=symbol, bg=bg, fg=fg, font=("Segoe UI", 9, "bold"), width=2, pady=2).pack(side="left", padx=(0, 6))
            tk.Label(item, text=name, bg=BG, fg=fg if active or done else MUT, font=("Segoe UI", 9, "bold" if active else "normal")).pack(side="left")

    def _clear(self):
        for w in self.content.winfo_children():
            w.destroy()
        for w in self.footer.winfo_children():
            w.destroy()

    def _goto(self, n: int):
        self._req_id += 1
        self.step = n
        self._update_progress()
        self._clear()
        [self._step0_greeting, self._step1_holiday, self._step2_links, self._step3_emoji, self._step4_message][n]()

    def _back(self):
        if self.step > 0:
            self._goto(self.step - 1)

    def _advance(self):
        if self.step < len(STEP_NAMES) - 1:
            self._goto(self.step + 1)

    def _content_card(self):
        card = make_card(self.content, padx=24, pady=20)
        card.pack(fill="both", expand=True, pady=(0, 8))
        return card

    def _make_scroll_area(self, parent, height=None):
        wrap = tk.Frame(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        wrap.pack(fill="both", expand=True, pady=(12, 0))
        canvas = tk.Canvas(wrap, bg=CARD, bd=0, highlightthickness=0, height=height)
        sb = tk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=CARD)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def resize_inner(event):
            canvas.itemconfigure(win_id, width=event.width)
        canvas.bind("<Configure>", resize_inner)
        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(-(e.delta // 120), "units"))
        return canvas, inner

    # ────────────────────────────────────────────────────────────────────────
    #  STEP 0
    # ────────────────────────────────────────────────────────────────────────
    def _step0_greeting(self):
        card = self._content_card()
        tk.Label(card, text="Good morning, Myke! 🌞", font=("Segoe UI", 20, "bold"), bg=CARD, fg=PRI).pack(anchor="w")
        tk.Label(card, text="Choose a greeting to start today's message.", font=FONT_BODY, bg=CARD, fg=SEC).pack(anchor="w", pady=(4, 16))

        greetings = self.cfg.get("greetings", []) or DEFAULT_CONFIG["greetings"]
        recent = self.cfg.get("recent_greetings", [])
        proj = self.cfg.get("project_name", "Team")
        day = self.now.strftime("%A")
        not_recent = [g for g in greetings if g not in recent] or greetings
        if not self.chosen_greeting.get() or self.chosen_greeting.get() not in greetings:
            self.chosen_greeting.set(not_recent[0])

        list_box = tk.Frame(card, bg=CARD)
        list_box.pack(fill="both", expand=True)
        for g in greetings:
            self._render_greeting_row(list_box, g, proj, day, recent)

        ghost_btn(self.footer, "⚙ Settings", self._open_settings).pack(side="left")
        primary_btn(self.footer, "Next  →", self._on_greeting_next).pack(side="right")

    def _render_greeting_row(self, parent, template, proj, day, recent):
        display = template.format(project=proj, day=day)
        row = tk.Frame(parent, bg=CARD_SOFT, padx=12, pady=10, highlightthickness=1, highlightbackground=BORDER)
        row.pack(fill="x", pady=5)
        rb = tk.Radiobutton(row, variable=self.chosen_greeting, value=template, bg=CARD_SOFT, activebackground=CARD_SOFT, selectcolor=CARD_SOFT, cursor="hand2")
        rb.pack(side="left")
        tk.Label(row, text=display, bg=CARD_SOFT, fg=PRI, font=FONT_BODY).pack(side="left", padx=(8, 0))
        if template in recent[-3:]:
            rounded_label(row, "used recently", ACCENT_SOFT, ACCENT).pack(side="right")
        for w in [row] + row.winfo_children():
            w.bind("<Button-1>", lambda e, t=template: self.chosen_greeting.set(t))

    def _on_greeting_next(self):
        g = self.chosen_greeting.get()
        if not g:
            messagebox.showwarning("Greeting", "Please choose a greeting.", parent=self.root)
            return
        recent = self.cfg.setdefault("recent_greetings", [])
        if g in recent:
            recent.remove(g)
        recent.append(g)
        self.cfg["recent_greetings"] = recent[-12:]
        save_config(self.cfg)
        self._advance()

    # ────────────────────────────────────────────────────────────────────────
    #  STEP 1
    # ────────────────────────────────────────────────────────────────────────
    def _step1_holiday(self):
        card = self._content_card()
        tk.Label(card, text="Choose today's occasion", font=FONT_H1, bg=CARD, fg=PRI).pack(anchor="w")
        self._hol_status = tk.Label(card, text="Fetching from configured API sources…", font=FONT_BODY, bg=CARD, fg=MUT)
        self._hol_status.pack(anchor="w", pady=(4, 0))
        self._hol_canvas, self._hol_list = self._make_scroll_area(card)
        self._holiday_var = tk.IntVar(value=-1)
        ghost_btn(self.footer, "← Back", self._back).pack(side="right", padx=(0, 10))
        primary_btn(self.footer, "Next  →", self._on_holiday_next).pack(side="right")
        ghost_btn(self.footer, "↺ Refresh", self._load_holidays).pack(side="left")
        self._load_holidays()

    def _load_holidays(self):
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

    def _populate_holidays(self, holidays):
        self.holidays = holidays
        self._holiday_var = tk.IntVar(value=-1)
        for w in self._hol_list.winfo_children():
            w.destroy()
        if not holidays:
            self._hol_status.config(text="No occasions found for today. You can continue and make a generic message.", fg=RED)
            return
        self._hol_status.config(text=f"{len(holidays)} occasions found today.", fg=GREEN)
        prev_idx = 0
        if self.chosen_holiday:
            for i, h in enumerate(holidays):
                if h["name"].lower() == self.chosen_holiday["name"].lower():
                    prev_idx = i
                    break
        self._holiday_var.set(prev_idx)
        self.chosen_holiday = holidays[prev_idx]
        for i, h in enumerate(holidays):
            self._render_holiday_row(i, h)

    def _render_holiday_row(self, idx, h):
        warn = h["sensitivity"] == "warn"
        rbg = WARN_BG if warn else CARD
        row = tk.Frame(self._hol_list, bg=rbg, padx=12, pady=10)
        row.pack(fill="x")
        rb = tk.Radiobutton(row, variable=self._holiday_var, value=idx, command=lambda i=idx: self._select_holiday(i), bg=rbg, activebackground=rbg, selectcolor=rbg, cursor="hand2")
        rb.pack(side="left")
        body = tk.Frame(row, bg=rbg)
        body.pack(side="left", fill="x", expand=True, padx=(8, 0))
        title = h["name"] + ("   ⚠️ handle with care" if warn else "")
        tk.Label(body, text=title, font=FONT_BODY, bg=rbg, fg=WARN_FG if warn else PRI, anchor="w").pack(anchor="w")
        score = fun_priority(h)
        source_text = f"Source: {h['source']}" + ("  •  fun pick" if score >= 80 else "")
        tk.Label(body, text=source_text, font=FONT_SMALL, bg=rbg, fg=MUT).pack(anchor="w")
        hline(self._hol_list)
        for w in [row, body] + body.winfo_children():
            w.bind("<Button-1>", lambda e, i=idx: (self._holiday_var.set(i), self._select_holiday(i)))

    def _select_holiday(self, idx):
        if 0 <= idx < len(self.holidays):
            self.chosen_holiday = self.holidays[idx]

    def _on_holiday_next(self):
        sel = self._holiday_var.get()
        self.chosen_holiday = self.holidays[sel] if 0 <= sel < len(self.holidays) else None
        self.chosen_link = None
        self.selected_emojis = []
        self._advance()

    # ────────────────────────────────────────────────────────────────────────
    #  STEP 2
    # ────────────────────────────────────────────────────────────────────────
    def _step2_links(self):
        card = self._content_card()
        name = self.chosen_holiday["name"] if self.chosen_holiday else "today"
        tk.Label(card, text="Choose a link to include", font=FONT_H1, bg=CARD, fg=PRI).pack(anchor="w")
        tk.Label(card, text=f"Finding links relevant to: {name}", font=FONT_BODY, bg=CARD, fg=SEC).pack(anchor="w", pady=(4, 0))
        self._link_status = tk.Label(card, text="🔍 Searching…", font=FONT_BODY, bg=CARD, fg=MUT)
        self._link_status.pack(anchor="w", pady=(12, 0))
        self._link_canvas, self._link_frm = self._make_scroll_area(card)
        self._link_var = tk.IntVar(value=-1)
        ghost_btn(self.footer, "← Back", self._back).pack(side="right", padx=(0, 10))
        primary_btn(self.footer, "Next  →", self._advance).pack(side="right")
        req_id = self._req_id
        def worker():
            lks = find_links(self.chosen_holiday) if self.chosen_holiday else []
            if self._req_id == req_id:
                self.root.after(0, lambda: self._populate_links(lks))
        threading.Thread(target=worker, daemon=True).start()

    def _populate_links(self, links):
        self.links = links
        self.chosen_link = None
        self._link_var = tk.IntVar(value=-1)
        for w in self._link_frm.winfo_children():
            w.destroy()
        self._link_status.config(text=f"{len(links)} links found." if links else "No links found. You can continue without one.", fg=GREEN if links else MUT)
        for i, lk in enumerate(links):
            self._render_link_row(i, lk)
        self._render_no_link_row()

    def _render_link_row(self, idx, lk):
        row = tk.Frame(self._link_frm, bg=CARD, padx=12, pady=10)
        row.pack(fill="x")
        rb = tk.Radiobutton(row, variable=self._link_var, value=idx, command=lambda i=idx: self._select_link(i), bg=CARD, activebackground=CARD, selectcolor=CARD, cursor="hand2")
        rb.pack(side="left", anchor="n", pady=(2, 0))
        body = tk.Frame(row, bg=CARD)
        body.pack(side="left", fill="x", expand=True, padx=(8, 0))
        tk.Label(body, text=f"{lk.get('icon', '🔗')}  {lk['label'][:82]}", font=FONT_BODY, bg=CARD, fg=PRI, wraplength=560, justify="left").pack(anchor="w")
        short = lk["url"][:90] + ("…" if len(lk["url"]) > 90 else "")
        url_lbl = tk.Label(body, text=short, font=FONT_MONO, bg=CARD, fg=LINK, cursor="hand2", wraplength=560, justify="left")
        url_lbl.pack(anchor="w", pady=(2, 0))
        url_lbl.bind("<Button-1>", lambda e, u=lk["url"]: webbrowser.open(u))
        hline(self._link_frm)
        for w in [row, body] + body.winfo_children():
            w.bind("<Button-1>", lambda e, i=idx: (self._link_var.set(i), self._select_link(i)))

    def _render_no_link_row(self):
        row = tk.Frame(self._link_frm, bg=CARD_SOFT, padx=12, pady=10)
        row.pack(fill="x")
        tk.Radiobutton(
            row,
            text="No link — message only",
            variable=self._link_var,
            value=-1,
            command=lambda: setattr(self, "chosen_link", None),
            font=FONT_BODY,
            bg=CARD_SOFT,
            fg=SEC,
            activebackground=CARD_SOFT,
            selectcolor=CARD_SOFT,
            cursor="hand2",
        ).pack(anchor="w")

    def _select_link(self, idx):
        if 0 <= idx < len(self.links):
            self.chosen_link = self.links[idx]

    # ────────────────────────────────────────────────────────────────────────
    #  STEP 3 EMOJI
    # ────────────────────────────────────────────────────────────────────────
    def _step3_emoji(self):
        card = self._content_card()
        name = self.chosen_holiday["name"] if self.chosen_holiday else self.now.strftime("%A")
        tk.Label(card, text="Add emojis", font=FONT_H1, bg=CARD, fg=PRI).pack(anchor="w")
        tk.Label(card, text=f"Choose one or more emojis for: {name}", font=FONT_BODY, bg=CARD, fg=SEC).pack(anchor="w", pady=(4, 12))

        selected_card = tk.Frame(card, bg=ACCENT_SOFT, padx=12, pady=8, highlightthickness=1, highlightbackground="#DCD3FF")
        selected_card.pack(fill="x", pady=(0, 12))
        self._selected_emoji_label = tk.Label(selected_card, text="Selected: None", bg=ACCENT_SOFT, fg=ACCENT, font=FONT_BODY)
        self._selected_emoji_label.pack(side="left")
        ghost_btn(selected_card, "Clear all", self._clear_emojis, font=FONT_SMALL, padx=8, pady=3).pack(side="right")

        self._emoji_grid = tk.Frame(card, bg=CARD)
        self._emoji_grid.pack(fill="both", expand=True)

        # No emoji option
        no_var = tk.BooleanVar(value=len(self.selected_emojis) == 0)
        self._no_emoji_var = no_var
        no_btn = tk.Checkbutton(
            self._emoji_grid,
            text="🚫  No emoji",
            variable=no_var,
            command=self._toggle_no_emoji,
            bg=CARD_SOFT,
            fg=SEC,
            font=FONT_BODY,
            activebackground=CARD_SOFT,
            selectcolor=CARD_SOFT,
            indicatoron=False,
            padx=12,
            pady=10,
            relief="flat",
            cursor="hand2",
        )
        no_btn.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        emojis = emoji_suggestions_for(name)
        self.emoji_vars = {}
        for i, emoji in enumerate(emojis, start=1):
            var = tk.BooleanVar(value=emoji in self.selected_emojis)
            self.emoji_vars[emoji] = var
            btn = tk.Checkbutton(
                self._emoji_grid,
                text=emoji,
                variable=var,
                command=lambda e=emoji: self._toggle_emoji(e),
                bg=CARD_SOFT,
                fg=PRI,
                font=FONT_EMOJI,
                activebackground=ACCENT_SOFT,
                selectcolor=ACCENT_SOFT,
                indicatoron=False,
                padx=12,
                pady=8,
                relief="flat",
                cursor="hand2",
            )
            r, c = divmod(i, 6)
            btn.grid(row=r, column=c, sticky="nsew", padx=5, pady=5)
        for c in range(6):
            self._emoji_grid.columnconfigure(c, weight=1)
        self._refresh_selected_emoji_label()

        ghost_btn(self.footer, "← Back", self._back).pack(side="right", padx=(0, 10))
        primary_btn(self.footer, "Next  →", self._advance).pack(side="right")

    def _toggle_no_emoji(self):
        if self._no_emoji_var.get():
            self.selected_emojis = []
            for var in self.emoji_vars.values():
                var.set(False)
        self._refresh_selected_emoji_label()

    def _toggle_emoji(self, emoji):
        if self.emoji_vars[emoji].get():
            if emoji not in self.selected_emojis:
                self.selected_emojis.append(emoji)
        else:
            if emoji in self.selected_emojis:
                self.selected_emojis.remove(emoji)
        self._no_emoji_var.set(len(self.selected_emojis) == 0)
        self._refresh_selected_emoji_label()

    def _clear_emojis(self):
        self.selected_emojis = []
        for var in self.emoji_vars.values():
            var.set(False)
        if hasattr(self, "_no_emoji_var"):
            self._no_emoji_var.set(True)
        self._refresh_selected_emoji_label()

    def _refresh_selected_emoji_label(self):
        text = "Selected: " + (" ".join(self.selected_emojis) if self.selected_emojis else "None")
        if hasattr(self, "_selected_emoji_label"):
            self._selected_emoji_label.config(text=text)

    # ────────────────────────────────────────────────────────────────────────
    #  STEP 4 MESSAGE
    # ────────────────────────────────────────────────────────────────────────
    def _step4_message(self):
        card = self._content_card()
        tk.Label(card, text="Your message is ready", font=FONT_H1, bg=CARD, fg=PRI).pack(anchor="w")
        tk.Label(card, text="Edit it if needed, then copy to Slack.", font=FONT_BODY, bg=CARD, fg=SEC).pack(anchor="w", pady=(4, 12))
        self.msg_txt = tk.Text(
            card,
            font=("Segoe UI", 12),
            height=11,
            relief="flat",
            wrap="word",
            bg="#FBFAFF",
            fg=PRI,
            insertbackground=PRI,
            padx=14,
            pady=14,
            highlightthickness=1,
            highlightbackground="#CEC6FF",
            highlightcolor=ACCENT,
        )
        self.msg_txt.pack(fill="both", expand=True)
        self.msg_txt.insert("1.0", self._build_message())
        self.char_label = tk.Label(card, text="", font=FONT_SMALL, bg=CARD, fg=MUT)
        self.char_label.pack(anchor="e", pady=(6, 0))
        self._update_char_count()
        self.msg_txt.bind("<KeyRelease>", lambda e: self._update_char_count())

        ghost_btn(self.footer, "← Back", self._back).pack(side="right", padx=(0, 10))
        self.copy_btn = primary_btn(self.footer, "📋 Copy message", self._copy)
        self.copy_btn.pack(side="right")
        ghost_btn(self.footer, "↺ Start over", lambda: self._goto(0)).pack(side="left")

    def _update_char_count(self):
        msg = self.msg_txt.get("1.0", "end").strip()
        self.char_label.config(text=f"{len(msg)} characters")

    def _build_message(self):
        proj = self.cfg.get("project_name", "Team")
        day = self.now.strftime("%A")
        greeting = (self.chosen_greeting.get() or "Good morning everyone").format(project=proj, day=day)
        if self.chosen_holiday:
            body = f"{greeting}, happy {self.chosen_holiday['name']}!"
        else:
            body = f"{greeting}, happy {day}!"
        if self.chosen_link:
            body += f"\n\n{self.chosen_link.get('icon', '🔗')}  {self.chosen_link['label']}\n{self.chosen_link['url']}"
        if self.selected_emojis:
            body += " " + " ".join(self.selected_emojis)
        return body

    def _copy(self):
        msg = self.msg_txt.get("1.0", "end").strip()
        self.root.clipboard_clear()
        self.root.clipboard_append(msg)
        self.root.update()
        self.copy_btn.config(text="✅ Copied!", bg=GREEN)
        if self._copy_job:
            self.root.after_cancel(self._copy_job)
        self._copy_job = self.root.after(2200, lambda: self.copy_btn.config(text="📋 Copy message", bg=ACCENT))

    def _open_settings(self):
        SettingsWindow(self.root, self.cfg, on_save=lambda: setattr(self, "cfg", load_config()))

# ─────────────────────────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
def main():
    root = tk.Tk()
    try:
        root.tk_setPalette(background=BG)
    except Exception:
        pass
    MorningMessageApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
