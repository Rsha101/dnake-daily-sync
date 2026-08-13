import os
import traceback
import json
import warnings
import datetime
import openpyxl
import requests
import urllib3
import zipfile
from flask import Flask
from apscheduler.schedulers.background import BackgroundScheduler

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = Flask(__name__)

# --- פונקציית הסנכרון המרכזית (אותה פונקציה בדיוק) ---
def run_sync_task():
    print("--- Starting Daily Sync Task ---")
    # ... כאן תדביק את כל הקוד שהיה לך בתוך ה-daily_sync פעם ...
    # (כדי לא להעמיס כאן, פשוט תדביק את הלוגיקה שלך בפנים)
    print("--- Sync Finished ---")

# --- הגדרת השעון הפנימי ---
scheduler = BackgroundScheduler()
# ירוץ כל יום בשעה 20:00 בדיוק
scheduler.add_job(func=run_sync_task, trigger='cron', hour=20, minute=0)
scheduler.start()

@app.route('/')
def home():
    return "Server is running and scheduler is active!"

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
