import pandas as pd
from datetime import datetime
import json

url = "https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOpQrStuVWxyz12345/export?format=csv"

try:
    df = pd.read_csv(url)
    print("✅ Loaded Google Sheet data:\n", df)

    # Convert DataFrame to list of dictionaries
    data = df.to_dict(orient="records")

    # Add timestamp
    result = {
        "last_updated": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
        "data": data
    }

    # Save as JSON
    with open("sheet_backup.json", "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=4)

    print("💾 Saved as sheet_backup.json")

except Exception as e:
    print("❌ Error:", e)
