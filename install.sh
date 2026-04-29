#!/bin/bash
# Holiday Morning Messenger — Linux Installer

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
    echo "  ERROR: Python 3.8+ not found."
    echo "  Install with:  sudo apt install python3  (Debian/Ubuntu)"
    echo "             or:  sudo dnf install python3  (Fedora)"
    exit 1
fi
echo "         $($PYTHON --version) — OK"

# ── 2. tkinter ────────────────────────────────────────────────────
echo "  [2/4]  Checking tkinter..."
if ! "$PYTHON" -c "import tkinter" 2>/dev/null; then
    echo "  tkinter not found. Attempting to install..."
    if command -v apt &>/dev/null; then
        sudo apt install -y python3-tk
    elif command -v dnf &>/dev/null; then
        sudo dnf install -y python3-tkinter
    elif command -v pacman &>/dev/null; then
        sudo pacman -S --noconfirm tk
    else
        echo "  ERROR: Could not install tkinter automatically."
        echo "  Install it manually for your distribution."
        exit 1
    fi
fi
echo "         tkinter — OK"

# ── 3. Desktop entry ──────────────────────────────────────────────
echo "  [3/4]  Creating desktop launcher..."
DESKTOP_DIR="$HOME/.local/share/applications"
DESKTOP_FILE="$DESKTOP_DIR/holiday-messenger.desktop"
PYTHON_PATH="$(command -v "$PYTHON")"

mkdir -p "$DESKTOP_DIR"
cat > "$DESKTOP_FILE" <<DESKEOF
[Desktop Entry]
Name=Holiday Morning Messenger
Comment=Daily team greeting generator
Exec=$PYTHON_PATH $SCRIPT
Icon=appointment-new
Terminal=false
Type=Application
Categories=Office;Utility;
DESKEOF
chmod +x "$DESKTOP_FILE"
update-desktop-database "$DESKTOP_DIR" 2>/dev/null || true
echo "         Desktop entry created — OK"

# ── 4. Autostart ──────────────────────────────────────────────────
echo "  [4/4]  Setting up startup launch..."
AUTOSTART_DIR="$HOME/.config/autostart"
mkdir -p "$AUTOSTART_DIR"
cp "$DESKTOP_FILE" "$AUTOSTART_DIR/holiday-messenger.desktop"
echo "         Autostart entry installed — OK"

# ── Done ──────────────────────────────────────────────────────────
echo ""
echo "  ============================================="
echo "   Installation complete!"
echo "  ============================================="
echo ""
echo "   Opens at login  :  Yes (~/.config/autostart)"
echo "   App launcher    :  Applications menu → Holiday Morning Messenger"
echo "   App folder      :  $DIR"
echo ""
read -rp "  Press Enter to launch the app now..."
"$PYTHON_PATH" "$SCRIPT" &
disown
