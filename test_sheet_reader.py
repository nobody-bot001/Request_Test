import pandas as pd
from datetime import datetime
import json
import io
import requests

SHEET_ID = "1PttAjbFtfvymn9pBpvZdMwNylzQNFVC9yGzvsMl2Vr8"
GID = "0"
csv_url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&gid={GID}"

try:
    r = requests.get(csv_url)
    r.raise_for_status()

    df = pd.read_csv(io.StringIO(r.text))
    print("✅ Loaded Google Sheet data:\n", df)

    data = df.to_dict(orient="records")
    result = {
        "last_updated": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
        "data": data
    }

    with open("sheet_backup.json", "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=4)
    print("💾 Saved as sheet_backup.json")

except Exception as e:
    print("❌ Error:", e)
