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

# העלמת אזהרות
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = Flask(__name__)

@app.route('/daily-sync')
def daily_sync():
    try:
        MONDAY_TOKEN = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjYyMzIxMjUxOSwiYWFpIjoxMSwidWlkIjo5NzYwMTM1NywiaWFkIjoiMjAyNi0wMi0xOVQwOTowODozMS4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MjkzNzUzNDYsInJnbiI6ImV1YzEifQ.EsITDIb08RaofyL9eIJae6eFJ_zBUiOBeCugjSMqoDE"
        BOARD_ID = "5091812790"
        USERNAME = "shahar_ro@mail.tel-aviv.gov.il"
        PASSWORD = "Rr304050!"
        
        output_dir = "temp_downloads"
        os.makedirs(output_dir, exist_ok=True)

        tz = datetime.timezone(datetime.timedelta(hours=3))
        now = datetime.datetime.now(tz)
        
        print("--- STARTING SYNC ---")
        print("1. Fetching existing items from Monday...")
        monday_url = "https://api.monday.com/v2"
        monday_headers = {"Authorization": MONDAY_TOKEN, "Content-Type": "application/json", "API-Version": "2023-10"}
        
        query_existing = '''
        query {
          boards(ids: %s) {
            items_page(limit: 500) {
              items {
                name
                column_values(ids: ["date_mm0kk5yt"]) {
                  text
                }
              }
            }
          }
        }
        ''' % BOARD_ID
        
        res_monday = requests.post(monday_url, json={"query": query_existing}, headers=monday_headers, verify=False)
        existing_records = set()
        
        if res_monday.status_code == 200:
            data = res_monday.json()
            items = data.get('data', {}).get('boards', [{}])[0].get('items_page', {}).get('items', [])
            for item in items:
                item_name = item.get('name', '').strip()
                cols = item.get('column_values', [])
                item_date = cols[0]['text'] if cols else ''
                if item_name and item_date:
                    existing_records.add(f"{item_name}_{item_date}")
        print(f"-> Found {len(existing_records)} existing records in Monday.")

        last_run_datetime = now - datetime.timedelta(days=2)
        start_ts = int(last_run_datetime.timestamp() * 1000)
        end_ts = int(now.timestamp() * 1000)

        print("2. Logging in to DNAKE API...")
        session = requests.Session()
        login_res = session.post(
            'https://eu-api-cloud.ss-iot.com/admin-api/system/auth/login', 
            headers={'Content-Type': 'application/json;charset=UTF-8', 'User-Agent': 'Mozilla/5.0'},
            json={"username": USERNAME, "password": PASSWORD},
            verify=False
        )
        
        token = login_res.json().get('data', {}).get('accessToken')
        if not token: 
            print("-> ERROR: Login failed!")
            raise Exception("Login failed")
        print("-> Logged in successfully!")

        print("3. Downloading Logs...")
        page_no = 1
        all_rows_content = []
        while True:
            export_res = session.get(
                'https://eu-api-cloud.ss-iot.com/admin-api/business/device-opendoor-log/exportDeviceOpendoorLogCsv',
                headers={'Authorization': f'Bearer {token}', 'Project-Id': '2051211421803474944', 'User-Agent': 'Mozilla/5.0'},
                params={'pageNo': str(page_no), 'pageSize': '1000', 'unlockTime[0]': start_ts, 'unlockTime[1]': end_ts, 'unlockWay': '1'},
                verify=False
            )
            print(f"-> Downloaded page {page_no}. File size: {len(export_res.content)} bytes.")
            temp_filename = os.path.join(output_dir, f"page_{page_no}.xlsx")
            with open(temp_filename, "wb") as f: f.write(export_res.content)
            
            try:
                wb = openpyxl.load_workbook(temp_filename)
                rows = list(wb.active.iter_rows(values_only=True))
                os.remove(temp_filename)
                print(f"-> Extracted {len(rows)} rows from page {page_no}.")
                if not rows: break
                if page_no == 1: all_rows_content.extend(rows)
                else: all_rows_content.extend(rows[1:])
                if len(rows) < 1000: break
                page_no += 1
            except Exception as ex:
                print(f"-> ERROR reading Excel file on page {page_no}: {str(ex)}")
                if os.path.exists(temp_filename): os.remove(temp_filename)
                break

        print(f"4. Syncing {len(all_rows_content)} rows to Monday...")
        sent_count = 0
        skipped_count = 0
        for row in all_rows_content[1:]:
            if not row or len(row) < 6 or row[5] is None: continue
            
            user_name = str(row[5]).strip()
            raw_date = row[0].strftime("%Y-%m-%d") if isinstance(row[0], datetime.datetime) else str(row[0]).split(' ')[0]
            
            if f"{user_name}_{raw_date}" in existing_records: 
                skipped_count += 1
                continue

            print(f"-> Sending new item: {user_name} ({raw_date})")
            query = 'mutation ($boardId: ID!, $itemName: String!, $columnValues: JSON!) { create_item (board_id: $boardId, item_name: $itemName, column_values: $columnValues) { id } }'
            vars = {"boardId": BOARD_ID, "itemName": user_name, "columnValues": json.dumps({"date_mm0kk5yt": {"date": raw_date}, "color_mm4nb4ob": {"label": "תחבר"}})}
            
            res = requests.post(monday_url, json={"query": query, "variables": vars}, headers=monday_headers, verify=False)
            if "errors" not in res.text:
                sent_count += 1
                existing_records.add(f"{user_name}_{raw_date}")
            else:
                print(f"-> ERROR from Monday for {user_name}: {res.text}")

        result_msg = f"Sync complete! Added {sent_count} items. Skipped {skipped_count} duplicates."
        print(result_msg)
        return result_msg, 200

    except Exception as e:
        error_msg = f"Error: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        return error_msg, 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
