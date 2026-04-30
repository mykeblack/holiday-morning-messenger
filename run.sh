#!/bin/bash
cd "$(dirname "$0")"

if command -v python3 >/dev/null 2>&1; then
  python3 holiday_messenger.py
elif command -v python >/dev/null 2>&1; then
  python holiday_messenger.py
else
  echo "Python was not found. Please install Python 3.10+."
  read -p "Press Enter to exit..."
  exit 1
fi

if [ $? -ne 0 ]; then
  echo
  read -p "Myke's morning message closed with an error. Press Enter to exit..."
fi
