import traceback
import json
import warnings
import datetime
import requests
import urllib3
from flask import Flask

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = Flask(__name__)

MONDAY_TOKEN = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjYyMzIxMjUxOSwiYWFpIjoxMSwidWlkIjo5NzYwMTM1NywiaWFkIjoiMjAyNi0wMi0xOVQwOTowODozMS4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MjkzNzUzNDYsInJnbiI6ImV1YzEifQ.EsITDIb08RaofyL9eIJae6eFJ_zBUiOBeCugjSMqoDE"
BOARD_LOGS = "5091812790"     
BOARD_STATS = "5091542066"    
USERNAME = "shahar_ro@mail.tel-aviv.gov.il"
PASSWORD = "Rr304050!"

def get_monday_headers():
    return {"Authorization": MONDAY_TOKEN, "Content-Type": "application/json", "API-Version": "2023-10"}

@app.route('/')
def home():
    return "Server is awake!", 200

@app.route('/morning-setup')
def morning_setup():
    try:
        tz = datetime.timezone(datetime.timedelta(hours=3))
        now = datetime.datetime.now(tz)
        item_name = now.strftime("%d/%m/%Y")
        
        monday_url = "https://api.monday.com/v2"
        query_check = 'query { boards(ids: %s) { items_page(limit: 100) { items { id name } } } }' % BOARD_STATS
        res_check = requests.post(monday_url, json={"query": query_check}, headers=get_monday_headers(), verify=False)
        items = res_check.json().get('data', {}).get('boards', [{}])[0].get('items_page', {}).get('items', [])
        
        for item in items:
            if item.get('name') == item_name:
                return f"Row '{item_name}' already exists in stats board.", 200

        query_create = 'mutation ($boardId: ID!, $itemName: String!) { create_item (board_id: $boardId, item_name: $itemName) { id } }'
        res_create = requests.post(monday_url, json={"query": query_create, "variables": {"boardId": BOARD_STATS, "itemName": item_name}}, headers=get_monday_headers(), verify=False)
        
        if "errors" not in res_create.text:
            return f"Morning setup complete. Created row '{item_name}'.", 200
        else:
            return "Failed to create morning row.", 500
    except Exception as e:
        return f"Error: {str(e)}", 500

@app.route('/daily-sync')
def daily_sync():
    try:
        tz = datetime.timezone(datetime.timedelta(hours=3))
        now = datetime.datetime.now(tz)
        
        today_name = now.strftime("%d/%m/%Y")
        today_date_excel_format = now.strftime("%Y-%m-%d")
        
        session = requests.Session()
        login_res = session.post(
            'https://eu-api-cloud.ss-iot.com/admin-api/system/auth/login', 
            headers={'Content-Type': 'application/json;charset=UTF-8', 'User-Agent': 'Mozilla/5.0'},
            json={"username": USERNAME, "password": PASSWORD},
            verify=False
        )
        token = login_res.json().get('data', {}).get('accessToken') or login_res.json().get('token')
        if not token: 
            return f"DEBUG ERROR: DNAKE Login failed. Response: {login_res.text}", 200

        all_records = []
        
        # מושכים נתונים ישירות מה-API החי (ללא אקסל) - 5 העמודים האחרונים (500 רשומות)
        for page_no in range(1, 6):
            page_res = session.get(
                'https://eu-api-cloud.ss-iot.com/admin-api/business/device-opendoor-log/page',
                headers={'Authorization': f'Bearer {token}', 'Project-Id': '2051211421803474944', 'User-Agent': 'Mozilla/5.0', 'Accept-Language': 'en-US'},
                params={'pageNo': str(page_no), 'pageSize': '100', 'unlockResult': '0'},
                verify=False
            )
            
            try:
                data = page_res.json()
                records = data.get('data', {})
                if isinstance(records, dict):
                    records = records.get('list', [])
                
                if not records:
                    break
                    
                all_records.extend(records)
                
                if len(records) < 100:
                    break # הגענו לעמוד האחרון
            except Exception as e:
                print(f"Failed to parse JSON on page {page_no}: {str(e)}")
                break

        monday_url = "https://api.monday.com/v2"
        seen_today = set() 
        sent_count = 0
        names_found_in_dnake_today = set()
        debug_dates_found_total = set()
        
        # עיבוד הנתונים
        for rec in all_records:
            ut = rec.get('unlockTime')
            user_name = rec.get('userName') or rec.get('personName') or rec.get('name')
            
            if not ut or not user_name:
                continue
                
            # חילוץ התאריך מתוך הנתון (יכול להיות טקסט או מספר)
            raw_date = None
            if isinstance(ut, str) and '-' in ut:
                raw_date = ut.split(' ')[0]
            elif isinstance(ut, (int, float)):
                if ut > 9999999999: # Timestamp במילישניות
                    dt = datetime.datetime.fromtimestamp(ut / 1000.0)
                else:
                    dt = datetime.datetime.fromtimestamp(ut)
                raw_date = dt.strftime('%Y-%m-%d')
                
            if not raw_date:
                continue
                
            debug_dates_found_total.add(raw_date)
            
            # מסננים רק את היום!
            if raw_date != today_date_excel_format:
                continue
                
            user_name = str(user_name).strip()
            names_found_in_dnake_today.add(user_name) 
            
            if user_name in seen_today: continue

            # שליחה למאנדיי (יומן כניסות)
            query_log = 'mutation ($boardId: ID!, $itemName: String!, $columnValues: JSON!) { create_item (board_id: $boardId, item_name: $itemName, column_values: $columnValues) { id } }'
            vars_log = {"boardId": BOARD_LOGS, "itemName": user_name, "columnValues": json.dumps({"date_mm0kk5yt": {"date": raw_date}, "color_mm4nb4ob": {"label": "תחבר"}})}
            
            res = requests.post(monday_url, json={"query": query_log, "variables": vars_log}, headers=get_monday_headers(), verify=False)
            if "errors" not in res.text:
                sent_count += 1
                seen_today.add(user_name) 
                    
        total_unique_visitors = len(seen_today)

        # עדכון סטטיסטיקה במאנדיי
        query_find = 'query { boards(ids: %s) { items_page(limit: 100) { items { id name } } } }' % BOARD_STATS
        res_find = requests.post(monday_url, json={"query": query_find}, headers=get_monday_headers(), verify=False)
        items = res_find.json().get('data', {}).get('boards', [{}])[0].get('items_page', {}).get('items', [])
        
        today_item_id = None
        for item in items:
            if item.get('name') == today_name:
                today_item_id = item.get('id')
                break
                
        if not today_item_id:
            query_create = 'mutation ($boardId: ID!, $itemName: String!) { create_item (board_id: $boardId, item_name: $itemName) { id } }'
            res_create = requests.post(monday_url, json={"query": query_create, "variables": {"boardId": BOARD_STATS, "itemName": today_name}}, headers=get_monday_headers(), verify=False)
            today_item_id = res_create.json().get('data', {}).get('create_item', {}).get('id')

        if today_item_id:
            query_update = 'mutation ($boardId: ID!, $itemId: ID!, $columnValues: JSON!) { change_multiple_column_values(board_id: $boardId, item_id: $itemId, column_values: $columnValues) { id } }'
            vars_update = {"boardId": BOARD_STATS, "itemId": today_item_id, "columnValues": json.dumps({"numeric_mm69bgft": str(total_unique_visitors)})}
            requests.post(monday_url, json={"query": query_update, "variables": vars_update}, headers=get_monday_headers(), verify=False)

        # דו"ח סיום
        result_msg = (
            f"Evening sync complete! Added {sent_count} names. Updated stats board with {total_unique_visitors} visitors.\n\n"
            f"--- DEBUG REPORT ---\n"
            f"API Endpoint Used: DIRECT JSON (Bye bye Excel!)\n"
            f"Total raw records fetched: {len(all_records)}\n"
            f"Dates found in recent records: {list(debug_dates_found_total)}\n"
            f"Names found specifically for today ({today_date_excel_format}): {list(names_found_in_dnake_today)}"
        )
        return result_msg, 200

    except Exception as e:
        error_msg = f"Error in daily sync: {str(e)}\n{traceback.format_exc()}"
        return error_msg, 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
