#!/usr/bin/env python3.11

import os
import shutil

# Ensure /tmp/balance_sheet_repo/ exists and sync data files there,
# since balancesheet.py hardcodes that path.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TMP_DIR = "/tmp/balance_sheet_repo"
os.makedirs(TMP_DIR, exist_ok=True)

for fname in ["balancesheet.json", "balancesheet.csv", "assets_db.csv"]:
    src = os.path.join(SCRIPT_DIR, fname)
    dst = os.path.join(TMP_DIR, fname)
    if os.path.exists(src):
        shutil.copy2(src, dst)

from balancesheet import AssetsManager

manager = AssetsManager()
manager.update_balance_sheet_db()
manager.update_legacy_assets_db()
manager.summarize_balance_sheet_db()
manager.generate_balance_sheet_chart()
manager.generate_legacy_assets_curve()
manager.show_assets()

# Copy updated files back to the repo directory
for fname in ["balancesheet.csv", "assets_db.csv", "balancesheet.png", "assets_curve.png"]:
    src = os.path.join(TMP_DIR, fname)
    dst = os.path.join(SCRIPT_DIR, fname)
    if os.path.exists(src):
        shutil.copy2(src, dst)
