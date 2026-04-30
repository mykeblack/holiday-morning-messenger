# 🌅 Myke’s Morning Message

A polished desktop app for creating cheerful morning Slack/messages with real holidays, fun novelty days, useful links, and emojis.

## ✨ Features

- Smart morning greeting picker
- Real-world holidays via Wikipedia and Nager.Date
- Fun novelty days via Checkiday, including things like Mr Potato Head Day
- Fun-first sorting so quirky days appear ahead of drier observances
- Auto-suggested links
- Multi-select emoji picker with relevant emoji suggestions
- One-click copy to Slack
- Custom app icon and branded colours
- Cross-platform Python app: Windows, macOS, Linux

## 🚀 Run from source

### Windows

Double-click:

```text
run.bat
```

Or run manually:

```powershell
python holiday_messenger.py
```

### macOS

```bash
chmod +x run.command
./run.command
```

### Linux

```bash
chmod +x run.sh
./run.sh
```

## 🐍 Requirements

- Python 3.10+
- Tkinter, usually included with Python on Windows and macOS

Test Tkinter with:

```bash
python -m tkinter
```

On Ubuntu/Debian, install Tkinter with:

```bash
sudo apt install python3-tk
```

## 🔑 Optional: enable Checkiday fun days

Checkiday is the best source for quirky novelty days.

1. Get an API key from the Checkiday API on APILayer.
2. Open the app.
3. Go to **Settings → API Sources**.
4. Enable **Checkiday – Fun Holidays**.
5. Paste your API key.

The app still works without Checkiday, but the fun-day list will be much better with it enabled.

## 📦 Build a Windows EXE

The easiest way to build the app into a Windows `.exe` is with PyInstaller.

### 1. Install PyInstaller

From the project folder:

```powershell
python -m pip install pyinstaller
```

If `pyinstaller` is not recognised as a command, use `python -m PyInstaller` in the commands below.

### 2. Make sure the assets exist

Your project should contain:

```text
holiday_messenger.py
assets\app_icon.ico
assets\app_icon.png
assets\app_logo.png
```

The `.ico` is used for the EXE file icon. The `.png` is used by Tkinter for the window/taskbar icon.

### 3. Clean old builds

If you have built before, clear the old PyInstaller files first:

```powershell
if (Test-Path build) { Remove-Item build -Recurse -Force }
if (Test-Path dist) { Remove-Item dist -Recurse -Force }
if (Test-Path holiday_messenger.spec) { Remove-Item holiday_messenger.spec -Force }
```

### 4. Build the EXE

```powershell
python -m PyInstaller --onefile --windowed `
  --icon "assets\app_icon.ico" `
  --add-data "assets\app_icon.png;assets" `
  --add-data "assets\app_logo.png;assets" `
  --name "MykesMorningMessage" `
  holiday_messenger.py
```

The finished app will be here:

```text
dist\MykesMorningMessage.exe
```

### 5. Debug build, if needed

If the app opens and closes immediately, build without `--windowed` so you can see errors:

```powershell
python -m PyInstaller --onefile `
  --icon "assets\app_icon.ico" `
  --add-data "assets\app_icon.png;assets" `
  --add-data "assets\app_logo.png;assets" `
  --name "MykesMorningMessage" `
  holiday_messenger.py
```

Then run:

```powershell
.\dist\MykesMorningMessage.exe
```

### Notes about config.json

The EXE creates `config.json` next to the EXE when it first runs. This is intentional so your settings and API keys remain editable after compiling.

## 🧪 Example message

```text
Morning all, happy National Pizza Day!

📖 Wikipedia: National Pizza Day
https://en.wikipedia.org/wiki/National_Pizza_Day

🍕 🔥 😄
```

## 👤 Author

Built by Myke Black.
