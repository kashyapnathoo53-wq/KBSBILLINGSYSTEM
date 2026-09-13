#!/bin/bash
cd "$(dirname "$0")"

echo "Installing/checking requirements (only needed the first time)..."
pip3 install -r requirements.txt >/dev/null 2>&1

echo "Starting the billing system..."
( sleep 2 && open http://127.0.0.1:5000 2>/dev/null || xdg-open http://127.0.0.1:5000 2>/dev/null ) &
python3 app.py
