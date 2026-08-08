import urllib.request
import urllib.parse
import json
from datetime import datetime

# Google Apps Script Web App URL
SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyOKWD7q_7YQF-TbeMSYRXksEubwHoWztZTnx8aIO4LOeWDmxKzz71UHt1HAs6z7fG_jw/exec"

def append_row(nama, kegiatan, nominal, bank, rekening, status, bukti_path):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Payload matching typical Google Apps Script expected format
    data = {
        'Tanggal': now,
        'Nama': nama,
        'Kegiatan': kegiatan,
        'Nominal': nominal,
        'Bank': bank,
        'Rekening': rekening,
        'Status': status,
        'Bukti_Path': bukti_path
    }
    
    try:
        # We try to send it as JSON first
        json_data = json.dumps(data).encode('utf-8')
        req = urllib.request.Request(SCRIPT_URL, data=json_data, headers={'Content-Type': 'application/json'})
        with urllib.request.urlopen(req) as response:
            result = response.read().decode('utf-8')
            print(f"Google Apps Script Response (JSON): {result}")
            return True
    except Exception as e:
        print(f"Failed sending JSON to Google Sheet, trying Form-Data... Error: {e}")
        try:
            # Fallback to URL-encoded form data
            encoded_data = urllib.parse.urlencode(data).encode('utf-8')
            req = urllib.request.Request(SCRIPT_URL, data=encoded_data)
            with urllib.request.urlopen(req) as response:
                result = response.read().decode('utf-8')
                print(f"Google Apps Script Response (Form): {result}")
                return True
        except Exception as e2:
            print(f"Failed sending Form-Data to Google Sheet. Error: {e2}")
            return False
