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
        query_check = 'query { boards(ids: %s) { items_page(limit: 50) { items { id name } } } }' % BOARD_STATS
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
        output_dir = "temp_downloads"
        os.makedirs(output_dir, exist_ok=True)

        tz = datetime.timezone(datetime.timedelta(hours=3))
        now = datetime.datetime.now(tz)
        
        today_name = now.strftime("%d/%m/%Y")
        today_date_excel_format = now.strftime("%Y-%m-%d")
        
        last_run_datetime = now - datetime.timedelta(days=2)
        start_ts = int(last_run_datetime.timestamp() * 1000)
        end_ts = int(now.timestamp() * 1000)
        
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

        page_no = 1
        all_rows_content = []
        
        # --- כאן מתבצעת הבדיקה הקריטית ---
        export_res = session.get(
            'https://eu-api-cloud.ss-iot.com/admin-api/business/device-opendoor-log/exportDeviceOpendoorLogCsv',
            headers={'Authorization': f'Bearer {token}', 'Project-Id': '2051211421803474944', 'User-Agent': 'Mozilla/5.0'},
            params={'pageNo': str(page_no), 'pageSize': '1000', 'unlockTime[0]': start_ts, 'unlockTime[1]': end_ts, 'unlockWay': '1'},
            verify=False
        )
        
        content_size = len(export_res.content)
        if content_size < 150 or "code" in export_res.text:
            return f"DEBUG ERROR: DNAKE API returned an error or empty data. Size: {content_size} bytes. Content: {export_res.text}", 200
        
        temp_filename = os.path.join(output_dir, f"page_{page_no}.xlsx")
        with open(temp_filename, "wb") as f: f.write(export_res.content)
        
        try:
            wb = openpyxl.load_workbook(temp_filename)
            rows = list(wb.active.iter_rows(values_only=True))
            os.remove(temp_filename)
            all_rows_content.extend(rows)
        except Exception as ex:
            if os.path.exists(temp_filename): os.remove(temp_filename)
            # אם הקובץ שבור, נדפיס את השגיאה ונראה מה באמת ירד!
            return f"DEBUG ERROR: Failed to open Excel file! Error: {str(ex)}. Downloaded size: {content_size} bytes. File start: {export_res.content[:50]}", 200
        # ------------------------------------

        monday_url = "https://api.monday.com/v2"
        seen_today = set() 
        sent_count = 0
        debug_dates_found = set()
        
        if len(all_rows_content) > 1:
            for row in all_rows_content[1:]:
                if not row or len(row) < 6 or row[5] is None: continue
                
                raw_date = row[0].strftime("%Y-%m-%d") if isinstance(row[0], datetime.datetime) else str(row[0]).split(' ')[0]
                debug_dates_found.add(raw_date) 
                
                if raw_date != today_date_excel_format:
                    continue
                
                user_name = str(row[5]).strip()
                if user_name in seen_today: continue

                query_log = 'mutation ($boardId: ID!, $itemName: String!, $columnValues: JSON!) { create_item (board_id: $boardId, item_name: $itemName, column_values: $columnValues) { id } }'
                vars_log = {"boardId": BOARD_LOGS, "itemName": user_name, "columnValues": json.dumps({"date_mm0kk5yt": {"date": raw_date}, "color_mm4nb4ob": {"label": "תחבר"}})}
                
                res = requests.post(monday_url, json={"query": query_log, "variables": vars_log}, headers=get_monday_headers(), verify=False)
                if "errors" not in res.text:
                    sent_count += 1
                    seen_today.add(user_name) 
                    
        total_unique_visitors = len(seen_today)
        
        if total_unique_visitors == 0:
            debug_msg = (f"DEBUG INFO: Expected date: {today_date_excel_format}. "
                         f"Total rows successfully extracted: {len(all_rows_content)}. "
                         f"Dates actually found in Excel: {list(debug_dates_found)}")
            return debug_msg, 200

        query_find = 'query { boards(ids: %s) { items_page(limit: 50) { items { id name } } } }' % BOARD_STATS
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

        result_msg = f"Evening sync complete! Added {sent_count} names. Updated stats board with {total_unique_visitors} visitors."
        return result_msg, 200

    except Exception as e:
        error_msg = f"Error in daily sync: {str(e)}\n{traceback.format_exc()}"
        return error_msg, 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
