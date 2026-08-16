import os
import traceback
import json
import warnings
import datetime
import openpyxl
import requests
import urllib3
from flask import Flask

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = Flask(__name__)

# --- משתני סביבה והגדרות ---
MONDAY_TOKEN = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjYyMzIxMjUxOSwiYWFpIjoxMSwidWlkIjo5NzYwMTM1NywiaWFkIjoiMjAyNi0wMi0xOVQwOTowODozMS4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MjkzNzUzNDYsInJnbiI6ImV1YzEifQ.EsITDIb08RaofyL9eIJae6eFJ_zBUiOBeCugjSMqoDE"
BOARD_LOGS = "5091812790"     # יומן כניסות (שמות)
BOARD_STATS = "5091542066"    # מעקב יזמים (סטטיסטיקה יומי)
USERNAME = "shahar_ro@mail.tel-aviv.gov.il"
PASSWORD = "Rr304050!"

def get_monday_headers():
    return {"Authorization": MONDAY_TOKEN, "Content-Type": "application/json", "API-Version": "2023-10"}

# --- 1. משימת שרת ער (Keep-alive) ---
@app.route('/')
def home():
    return "Server is awake!", 200

# --- 2. משימת בוקר (08:00): פתיחת שורה ריקה ליום החדש ---
@app.route('/morning-setup')
def morning_setup():
    try:
        tz = datetime.timezone(datetime.timedelta(hours=3))
        now = datetime.datetime.now(tz)
        
        # השם של השורה יהיה פשוט היום בחודש (למשל "16" או "17") כדי שיתאים למה שאתם רגילים
        item_name = str(now.day) 
        
        print(f"--- STARTING MORNING SETUP FOR DAY {item_name} ---")
        monday_url = "https://api.monday.com/v2"
        
        # בדיקה האם כבר קיימת שורה עם התאריך של היום (למקרה שהרצנו פעמיים)
        query_check = '''
        query {
          boards(ids: %s) {
            items_page(limit: 20) {
              items { id name }
            }
          }
        }
        ''' % BOARD_STATS
        
        res_check = requests.post(monday_url, json={"query": query_check}, headers=get_monday_headers(), verify=False)
        items = res_check.json().get('data', {}).get('boards', [{}])[0].get('items_page', {}).get('items', [])
        
        for item in items:
            if item.get('name') == item_name:
                print("-> Row already exists for today. Skipping creation.")
                return f"Row '{item_name}' already exists in stats board.", 200

        # אם השורה לא קיימת - יוצרים אותה
        print("-> Creating new row for today...")
        query_create = 'mutation ($boardId: ID!, $itemName: String!) { create_item (board_id: $boardId, item_name: $itemName) { id } }'
        vars_create = {"boardId": BOARD_STATS, "itemName": item_name}
        
        res_create = requests.post(monday_url, json={"query": query_create, "variables": vars_create}, headers=get_monday_headers(), verify=False)
        
        if "errors" not in res_create.text:
            print("-> Morning row created successfully!")
            return f"Morning setup complete. Created row '{item_name}'.", 200
        else:
            print(f"-> ERROR from Monday: {res_create.text}")
            return "Failed to create morning row.", 500

    except Exception as e:
        error_msg = f"Error in morning setup: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        return error_msg, 500


# --- 3. משימת ערב (20:00): סנכרון DNAKE ועדכון שני הלוחות ---
@app.route('/daily-sync')
def daily_sync():
    try:
        output_dir = "temp_downloads"
        os.makedirs(output_dir, exist_ok=True)

        tz = datetime.timezone(datetime.timedelta(hours=3))
        now = datetime.datetime.now(tz)
        today_name = str(now.day)
        
        start_of_day = now.replace(hour=0, minute=0, second=0, microsecond=0)
        start_ts = int(start_of_day.timestamp() * 1000)
        end_ts = int(now.timestamp() * 1000)
        
        print("--- STARTING DAILY SYNC (EVENING) ---")
        
        # שלב א': התחברות ל-DNAKE
        print("1. Logging in to DNAKE API...")
        session = requests.Session()
        login_res = session.post(
            'https://eu-api-cloud.ss-iot.com/admin-api/system/auth/login', 
            headers={'Content-Type': 'application/json;charset=UTF-8', 'User-Agent': 'Mozilla/5.0'},
            json={"username": USERNAME, "password": PASSWORD},
            verify=False
        )
        login_data = login_res.json()
        token = login_data.get('data', {}).get('accessToken') or login_data.get('token')
        if not token: raise Exception("DNAKE Login failed")
        print("-> Logged in successfully!")

        # שלב ב': הורדת הלוגים של היום בלבד
        print("2. Downloading Today's Logs...")
        page_no = 1
        all_rows_content = []
        while True:
            export_res = session.get(
                'https://eu-api-cloud.ss-iot.com/admin-api/business/device-opendoor-log/exportDeviceOpendoorLogCsv',
                headers={'Authorization': f'Bearer {token}', 'Project-Id': '2051211421803474944', 'User-Agent': 'Mozilla/5.0'},
                params={'pageNo': str(page_no), 'pageSize': '1000', 'unlockTime[0]': start_ts, 'unlockTime[1]': end_ts, 'unlockWay': '1'},
                verify=False
            )
            content_size = len(export_res.content)
            if content_size < 500: break # קובץ ריק או שגיאה
            
            temp_filename = os.path.join(output_dir, f"page_{page_no}.xlsx")
            with open(temp_filename, "wb") as f: f.write(export_res.content)
            
            try:
                wb = openpyxl.load_workbook(temp_filename)
                rows = list(wb.active.iter_rows(values_only=True))
                os.remove(temp_filename)
                if not rows: break
                if page_no == 1: all_rows_content.extend(rows)
                else: all_rows_content.extend(rows[1:])
                if len(rows) < 1000: break
                page_no += 1
            except Exception:
                if os.path.exists(temp_filename): os.remove(temp_filename)
                break

        # שלב ג': שליחת שמות ליומן כניסות + ספירה מקומית
        print(f"3. Syncing {len(all_rows_content)} rows to Monday (Logs Board)...")
        monday_url = "https://api.monday.com/v2"
        seen_today = set() 
        sent_count = 0
        
        if len(all_rows_content) > 1:
            for row in all_rows_content[1:]:
                if not row or len(row) < 6 or row[5] is None: continue
                
                user_name = str(row[5]).strip()
                raw_date = row[0].strftime("%Y-%m-%d") if isinstance(row[0], datetime.datetime) else str(row[0]).split(' ')[0]
                
                if user_name in seen_today: continue

                query_log = 'mutation ($boardId: ID!, $itemName: String!, $columnValues: JSON!) { create_item (board_id: $boardId, item_name: $itemName, column_values: $columnValues) { id } }'
                vars_log = {"boardId": BOARD_LOGS, "itemName": user_name, "columnValues": json.dumps({"date_mm0kk5yt": {"date": raw_date}, "color_mm4nb4ob": {"label": "תחבר"}})}
                
                res = requests.post(monday_url, json={"query": query_log, "variables": vars_log}, headers=get_monday_headers(), verify=False)
                if "errors" not in res.text:
                    sent_count += 1
                    seen_today.add(user_name) 
                    
        total_unique_visitors = len(seen_today)
        print(f"-> Sent {sent_count} names to Logs Board. Total Unique Visitors: {total_unique_visitors}")

        # שלב ד': עדכון מספר היזמים בלוח הסטטיסטיקה של הבוקר
        print("4. Updating Daily Stats Board...")
        
        # מחפשים את השורה של היום
        query_find = '''
        query { boards(ids: %s) { items_page(limit: 15) { items { id name } } } }
        ''' % BOARD_STATS
        res_find = requests.post(monday_url, json={"query": query_find}, headers=get_monday_headers(), verify=False)
        items = res_find.json().get('data', {}).get('boards', [{}])[0].get('items_page', {}).get('items', [])
        
        today_item_id = None
        for item in items:
            if item.get('name') == today_name:
                today_item_id = item.get('id')
                break
                
        # אם מישהו בטעות מחק את השורה של הבוקר, ניצור אותה מחדש עכשיו
        if not today_item_id:
            print("-> Morning row not found! Creating it now...")
            query_create = 'mutation ($boardId: ID!, $itemName: String!) { create_item (board_id: $boardId, item_name: $itemName) { id } }'
            res_create = requests.post(monday_url, json={"query": query_create, "variables": {"boardId": BOARD_STATS, "itemName": today_name}}, headers=get_monday_headers(), verify=False)
            today_item_id = res_create.json().get('data', {}).get('create_item', {}).get('id')

        # עדכון העמודה המספרית
        if today_item_id:
            query_update = 'mutation ($boardId: ID!, $itemId: ID!, $columnValues: JSON!) { change_multiple_column_values(board_id: $boardId, item_id: $itemId, column_values: $columnValues) { id } }'
            vars_update = {
                "boardId": BOARD_STATS, 
                "itemId": today_item_id, 
                "columnValues": json.dumps({"numeric_mm69bgft": total_unique_visitors})
            }
            res_update = requests.post(monday_url, json={"query": query_update, "variables": vars_update}, headers=get_monday_headers(), verify=False)
            
            if "errors" not in res_update.text:
                print(f"-> Successfully updated item {today_item_id} with count {total_unique_visitors}!")
            else:
                print(f"-> ERROR updating count: {res_update.text}")

        result_msg = f"Evening sync complete! Added {sent_count} names to logs. Updated stats board with {total_unique_visitors} visitors."
        return result_msg, 200

    except Exception as e:
        error_msg = f"Error in daily sync: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        return error_msg, 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
