# Holiday Morning Messenger v2 🌅

A 3-step wizard desktop app that opens on startup, fetches today's
holidays from live APIs, and builds a copy-ready Slack team message.

---

## Quick start

```bash
python3 holiday_messenger.py
```

No extra packages needed — uses Python's built-in `urllib` and `tkinter` only.

---

## The 3-step wizard

| Step | What you do |
|------|-------------|
| **1 – Greeting** | Pick a greeting opener. Recently used ones are labelled so you don't repeat yourself. |
| **2 – Occasion** | Browse today's holidays/observances fetched live from your API sources. Culturally sensitive days are flagged ⚠️ or filtered out entirely. |
| **3 – Link** | Choose a link to append (Wikipedia, official site, related pages) — or pick "no link". |

Then edit the final message if needed and click **📋 Copy** to paste into Slack.

---

## API sources

Two sources are enabled by default (no key required):

| Source | Type | Key needed? |
|--------|------|-------------|
| **Wikipedia** | Observances & fun days extracted from Wikipedia date pages | No |
| **Nager.Date** | Official public holidays by country | No |
| **Calendarific** | Comprehensive worldwide holidays + observances | Yes (free) |
| **Abstract API** | Holiday data by country | Yes (free) |

### Getting free API keys (optional but recommended)

**Calendarific** — covers 230+ countries and includes many observances  
1. Sign up at https://calendarific.com  
2. Copy your API key  
3. Open the app → ⚙ → API Sources → paste the key → enable it  

**Abstract API**  
1. Sign up at https://app.abstractapi.com/api/holidays  
2. Copy your API key  
3. Same steps as above  

### Changing your country

Edit `config.json` and change the `"country"` field to your ISO 2-letter
country code (e.g. `"US"`, `"AU"`, `"DE"`, `"FR"`, `"CA"`).

---

## Sensitivity filter

When enabled (default), the filter:

- **Hard-blocks** days where saying "Happy X" would be deeply inappropriate,
  e.g. Holocaust Remembrance Day, Hiroshima Day — these are silently removed.
- **Flags with ⚠️** days like Cancer Day, Remembrance Day, etc. — shown with
  a warning so you can decide whether to use them.

Toggle it in ⚙ → General.

---

## Customising greetings

Open ⚙ → Greetings to add, edit, or delete greeting templates.

Available placeholders:

| Placeholder | Example output |
|-------------|----------------|
| `{project}` | Dev Team |
| `{day}` | Wednesday |

Example:
```
Morning {project} engineers, happy {day}!
```

---

## Setting up startup (opens automatically at login)

### Built-in (recommended)
1. Open the app → ⚙ → Startup
2. Click **Install startup launcher**

### Manual — Windows
1. Press `Win+R`, type `shell:startup`, press Enter
2. Create `holiday_messenger.bat` containing:
   ```batch
   @echo off
   start "" pythonw "C:\path\to\holiday_messenger.py"
   ```

### Manual — macOS
1. System Settings → General → Login Items → click +
2. Navigate to `holiday_messenger.py` and add it

### Manual — Linux
Add to GNOME Startup Applications or your `.bashrc`:
```bash
python3 /path/to/holiday_messenger.py &
```

---

## Files

```
holiday_messenger/
├── holiday_messenger.py   ← the app
├── config.json            ← your settings (edit freely)
└── README.md
```

Settings (project name, greetings, API keys, recent greetings)
are stored in `config.json` in the same folder. You can edit it
directly in any text editor.
