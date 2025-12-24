import pandas as pd
import requests
from datetime import date, timedelta
import json

today = date.today()
yesterday = today - timedelta(days=1)

file_path = "Desiredoutput_final_new.xlsx"
df = pd.read_excel(file_path, engine="openpyxl")

mapped_data = []

for _, row in df.iterrows():
    datas = {
        "company": "RAGA TEX INDIA PRIVATE LIMITED",
        "employee": row["Employee Code"],
        "time_logs": [
            {
                "activity_type": "Working Time",
                "from_time": yesterday.strftime("%Y-%m-%d"),
                "to_time": yesterday.strftime("%Y-%m-%d"),
                "hours": row["Total Hours"]
            }
        ]
    }
    mapped_data.append(datas)

# ✅ Print ALL data
print(mapped_data)

url = "http://192.168.1.208/api/resource/Timesheet"
headers = {
    "Authorization": "Bearer YOUR_API_KEY:YOUR_API_SECRET",
    "Content-Type": "application/json",
    "Accept": "application/json"
}

for record in mapped_data:
    response = requests.post(url, headers=headers, json=record)

    if response.status_code in (200, 201):
        print("Timesheet created successfully")
        print(response.json())
    else:
        print(f"Failed. Status: {response.status_code}")
        print(response.text)