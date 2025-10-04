import pandas as pd
from datetime import datetime, timedelta, timezone
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

    utc_time = datetime.now(timezone.utc)
    local_time = utc_time + timedelta(hours=5)  # Pakistan Standard Time (UTC+5)

    result = {
        "last_updated": utc_time.strftime("%Y-%m-%d %H:%M:%S UTC"),
        "local_time": local_time.strftime("%Y-%m-%d %I:%M:%S %p PKT"),
        "data": data
    }

    with open("sheet_backup.json", "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=4)

    print("💾 Saved as sheet_backup.json with local time included")

except Exception as e:
    print("❌ Error:", e)
