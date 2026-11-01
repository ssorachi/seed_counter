@echo off
rem 種子カウンター起動（初回は必要なライブラリを自動で入れます）
cd /d "%~dp0"
python -c "import cv2, scipy, skimage, openpyxl, pillow_heif" 2>nul || python -m pip install -r requirements.txt
python seed_counter_app.py
if errorlevel 1 pause
