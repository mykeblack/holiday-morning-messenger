#!/bin/bash
# Holiday Morning Messenger — macOS Installer
# Double-click this file in Finder to run.

set -euo pipefail
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
SCRIPT="$DIR/holiday_messenger.py"

echo ""
echo "  ============================================="
echo "   Holiday Morning Messenger  |  Installer"
echo "  ============================================="
echo ""

# ── 1. Python ─────────────────────────────────────────────────────
echo "  [1/4]  Checking Python..."

PYTHON=""
for cmd in python3 python; do
    if command -v "$cmd" &>/dev/null; then
        VER=$("$cmd" -c "import sys; print(sys.version_info.major * 10 + sys.version_info.minor)" 2>/dev/null || echo 0)
        if [ "$VER" -ge 38 ]; then
            PYTHON="$cmd"
            break
        fi
    fi
done

if [ -z "$PYTHON" ]; then
    echo ""
    echo "  ERROR: Python 3.8+ not found."
    echo ""
    echo "  Install it from https://www.python.org/downloads/"
    echo "  (The python.org installer includes tkinter — Homebrew's does not.)"
    echo ""
    read -rp "  Press Enter to exit..."
    exit 1
fi
echo "         $($PYTHON --version) — OK"

# ── 2. tkinter ────────────────────────────────────────────────────
echo "  [2/4]  Checking tkinter..."
if ! "$PYTHON" -c "import tkinter" 2>/dev/null; then
    echo ""
    echo "  ERROR: tkinter not available."
    echo ""
    echo "  If you installed Python via Homebrew, run:"
    echo "    brew install python-tk"
    echo "  Or reinstall Python from https://www.python.org/downloads/"
    echo ""
    read -rp "  Press Enter to exit..."
    exit 1
fi
echo "         tkinter — OK"

# ── 3. LaunchAgent (startup) ──────────────────────────────────────
echo "  [3/4]  Setting up startup launch..."
PLIST_DIR="$HOME/Library/LaunchAgents"
PLIST="$PLIST_DIR/com.holidaymessenger.app.plist"
PYTHON_PATH="$(command -v "$PYTHON")"

mkdir -p "$PLIST_DIR"
cat > "$PLIST" <<PLISTEOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>com.holidaymessenger.app</string>
  <key>ProgramArguments</key>
  <array>
    <string>$PYTHON_PATH</string>
    <string>$SCRIPT</string>
  </array>
  <key>RunAtLoad</key>
  <true/>
</dict>
</plist>
PLISTEOF

launchctl unload "$PLIST" 2>/dev/null || true
launchctl load  "$PLIST" 2>/dev/null && \
    echo "         LaunchAgent installed — app will open at each login" || \
    echo "         LaunchAgent created (will activate on next login)"

# ── 4. Dock / Applications alias ─────────────────────────────────
echo "  [4/4]  Creating app launcher..."
LAUNCHER="$HOME/Applications/Holiday Messenger.command"
mkdir -p "$HOME/Applications"
cat > "$LAUNCHER" <<LAUNCHEOF
#!/bin/bash
cd "$DIR"
exec "$PYTHON_PATH" "$SCRIPT"
LAUNCHEOF
chmod +x "$LAUNCHER"
echo "         Launcher: ~/Applications/Holiday Messenger.command"

# ── Done ──────────────────────────────────────────────────────────
echo ""
echo "  ============================================="
echo "   Installation complete!"
echo "  ============================================="
echo ""
echo "   Opens at login  :  Yes (LaunchAgent)"
echo "   Manual launcher :  ~/Applications/Holiday Messenger.command"
echo "   App folder      :  $DIR"
echo ""
read -rp "  Press Enter to launch the app now..."
"$PYTHON_PATH" "$SCRIPT" &
disown
