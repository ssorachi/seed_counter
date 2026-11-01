#!/bin/bash
# 種子カウンター起動（Mac）
cd "$(dirname "$0")"
python3 -c "import cv2, scipy, skimage, openpyxl, pillow_heif" 2>/dev/null || python3 -m pip install -r requirements.txt
python3 seed_counter_app.py
