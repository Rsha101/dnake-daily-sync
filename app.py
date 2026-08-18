import traceback
import json
import warnings
import datetime
import requests
import urllib3
from flask import Flask

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = Flask(__name__)

USERNAME = "shahar_ro@mail.tel-aviv.gov.il"
PASSWORD = "Rr304050!"

@app.route('/')
def home():
    return "Server is awake!", 200

@app.route('/morning-setup')
def morning_setup():
    return "Skipped for now.", 200

@app.route('/daily-sync')
def daily_sync():
    try:
        tz = datetime.timezone(datetime.timedelta(hours=3))
        now = datetime.datetime.now(tz)
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
            return f"DEBUG ERROR: DNAKE Login failed.", 200

        dnake_headers = {
            'Authorization': f'Bearer {token}', 
            'Project-Id': '2051211421803474944', 
            'Role-Type': '14',
            'Accept-Language': 'en_US',
            'Accept': 'application/json, text/plain, */*',
            'Origin': 'https://eu-cloud.dnake.com',
            'Referer': 'https://eu-cloud.dnake.com/',
            'User-Agent': 'Mozilla/5.0'
        }
        
        found_today = []

        # סורקים את כל שיטות הפתיחה מ-1 עד 10
        for way in range(1, 11):
            way_str = str(way)
            for page_no in range(1, 6):
                page_res = session.get(
                    'https://eu-api-cloud.ss-iot.com/admin-api/business/device-opendoor-log/page',
                    headers=dnake_headers,
                    params={'pageNo': str(page_no), 'pageSize': '100', 'unlockResult': '0', 'unlockWay': way_str},
                    verify=False
                )
                try:
                    data = page_res.json()
                    records = data.get('data', {})
                    if isinstance(records, dict):
                        records = records.get('list', [])
                    
                    if not records:
                        break
                        
                    for rec in records:
                        ut = rec.get('unlockTime')
                        user_name = rec.get('unlockUserName') or rec.get('userName') or rec.get('personName') or rec.get('name')
                        
                        if not ut or not user_name: continue
                            
                        raw_date = None
                        if isinstance(ut, str) and '-' in ut:
                            raw_date = ut.split(' ')[0]
                        elif isinstance(ut, (int, float)):
                            if ut > 9999999999:
                                dt = datetime.datetime.fromtimestamp(ut / 1000.0)
                            else:
                                dt = datetime.datetime.fromtimestamp(ut)
                            raw_date = dt.strftime('%Y-%m-%d')
                            
                        # אם זו כניסה מהיום - נשמור את השם והשיטה
                        if raw_date == today_date_excel_format:
                            found_today.append(f"{user_name} (Method: {way_str})")
                            
                    if len(records) < 100:
                        break
                except Exception as e:
                    break

        unique_found = list(set(found_today))

        result_msg = (
            f"--- DEBUG: QR CODE HUNTER ---\n"
            f"Date being checked: {today_date_excel_format}\n"
            f"People found today across ALL methods (1 to 10):\n"
            f"{json.dumps(unique_found, ensure_ascii=False, indent=2)}\n\n"
            f"Please tell me the 'Method' number of the QR user!"
        )
        return result_msg, 200

    except Exception as e:
        return f"Error: {str(e)}\n{traceback.format_exc()}", 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
