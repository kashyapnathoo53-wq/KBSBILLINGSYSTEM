@echo off
title Sri Radhe Enterprises - Billing System
cd /d "%~dp0"

echo Installing/checking requirements (only needed the first time)...
pip install -r requirements.txt >nul 2>&1

echo Starting the billing system...
start "" http://127.0.0.1:5000
python app.py

pause
