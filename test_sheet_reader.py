import pandas as pd
import json, io, requests, traceback, subprocess
from datetime import datetime, timedelta, timezone
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

def safe_serialize(obj):
    if obj is None: return None
    if isinstance(obj, (str, int, float, bool)): return obj
    if hasattr(obj, 'rgb') and obj.rgb: return str(obj.rgb)
    if hasattr(obj, '__str__'): return str(obj)
    return None

def get_cell_styles():
    return {
        "font": {"name": None, "size": None, "bold": None, "italic": None, "color": None},
        "fill": {"fgColor": None},
        "alignment": {"horizontal": None, "vertical": None, "wrap_text": None}
    }

class SafeJSONEncoder(json.JSONEncoder):
    def default(self, obj):
        if hasattr(obj, '__str__'): return str(obj)
        return super().default(obj)

try:
    print("📥 Fetching Google Sheet as CSV...")

    SHEET_ID = "1cmDXt7UTIKBVXBHhtZ0E4qMnJrRoexl2GmDFfTBl0Z4"
    GID = "1882612924"
    csv_url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&gid={GID}"

    r = requests.get(csv_url)
    r.raise_for_status()

    df = pd.read_csv(io.StringIO(r.text))
    print("✅ Loaded Google Sheet data")

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    for r_idx, row in enumerate(df.itertuples(index=False), 1):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    structured_timetable = {}
    sheet_name = ws.title
    structured_timetable[sheet_name] = {"rooms": {}, "time_slots": [], "all_cells": []}

    header_rows = [1]
    time_slot_maps = {1: {col: ws.cell(row=1, column=col).value for col in range(2, ws.max_column + 1)}}

    for slot_map in time_slot_maps.values():
        structured_timetable[sheet_name]["time_slots"].extend(slot_map.values())

    rooms = {}
    for row in range(2, ws.max_row + 1):
        cell = ws.cell(row=row, column=1)
        if cell.value:
            room_name = safe_serialize(cell.value)
            rooms[row] = room_name
            structured_timetable[sheet_name]["rooms"][room_name] = {"time_slots": {}, "row": row, "schedule": []}

    for row in range(1, ws.max_row + 1):
        time_slots = time_slot_maps[1]
        row_data = []
        for col in range(1, ws.max_column + 1):
            value = safe_serialize(ws.cell(row=row, column=col).value)
            cell_info = {
                "cell_reference": f"{get_column_letter(col)}{row}",
                "row": row,
                "column": col,
                "column_letter": get_column_letter(col),
                "value": value,
                "time_slot": time_slots.get(col),
                "room": rooms.get(row),
                "styles": get_cell_styles()
            }
            row_data.append(cell_info)

            if row > 1 and col > 1 and value:
                room_name = rooms.get(row)
                time_slot = time_slots.get(col)
                if room_name and value:
                    if time_slot not in structured_timetable[sheet_name]["rooms"][room_name]["time_slots"]:
                        structured_timetable[sheet_name]["rooms"][room_name]["time_slots"][time_slot] = []
                    class_info = {
                        "subject": value,
                        "time_slot": time_slot,
                        "cell_reference": f"{get_column_letter(col)}{row}",
                        "row": row,
                        "column": col,
                        "styles": get_cell_styles()
                    }
                    structured_timetable[sheet_name]["rooms"][room_name]["time_slots"][time_slot].append(class_info)
                    structured_timetable[sheet_name]["rooms"][room_name]["schedule"].append(class_info)
        structured_timetable[sheet_name]["all_cells"].append(row_data)

    with open("timetable_detailed.json", "w", encoding="utf-8") as f:
        json.dump(structured_timetable, f, indent=2, ensure_ascii=False, cls=SafeJSONEncoder)
    print("✅ Saved timetable_detailed.json")

    simplified_timetable = {}
    simplified_timetable[sheet_name] = []

    for room_name, room_data in structured_timetable[sheet_name]["rooms"].items():
        room_schedule = {"room": room_name, "schedule": []}
        for class_info in room_data["schedule"]:
            schedule_entry = {
                "time_slot": class_info["time_slot"],
                "subject": class_info["subject"],
                "location": f"Cell {class_info['cell_reference']}",
                "row": class_info["row"],
                "column": class_info["column"]
            }
            room_schedule["schedule"].append(schedule_entry)
        room_schedule["schedule"].sort(key=lambda x: x["column"])
        simplified_timetable[sheet_name].append(room_schedule)

    with open("timetable_simplified_1.json", "w", encoding="utf-8") as f:
        json.dump(simplified_timetable, f, indent=2, ensure_ascii=False, cls=SafeJSONEncoder)
    print("✅ Saved timetable_simplified_1.json")

    utc_time = datetime.now(timezone.utc)
    local_time = utc_time + timedelta(hours=5)
    status = {
        "status": "success",
        "message": "Timetable extraction completed successfully.",
        "last_updated_utc": utc_time.strftime("%Y-%m-%d %H:%M:%S UTC"),
        "local_time_pkt": local_time.strftime("%Y-%m-%d %I:%M:%S %p PKT")
    }

except Exception as e:
    print("❌ Error:", e)
    status = {
        "status": "error",
        "message": str(e),
        "traceback": traceback.format_exc()
    }

with open("status.json", "w", encoding="utf-8") as f:
    json.dump(status, f, indent=2, ensure_ascii=False)
print("💾 Saved status.json")

# ---------------------------------------------------------------------
# STEP 4: PUSH TO GITHUB
# ---------------------------------------------------------------------
try:
    print("🚀 Pushing JSON files to GitHub...")

    subprocess.run(["git", "config", "--global", "user.name", "github-actions[bot]"], check=True)
    subprocess.run(["git", "config", "--global", "user.email", "github-actions[bot]@users.noreply.github.com"], check=True)

    subprocess.run(["git", "pull", "--rebase", "origin", "main"], check=False)
    subprocess.run(["git", "add", "timetable_detailed.json", "timetable_simplified_1.json", "status.json"], check=True)
    subprocess.run(["git", "commit", "-m", "🔄 Auto-update timetable JSONs from Google Sheet"], check=False)
    subprocess.run(["git", "push", "origin", "main"], check=False)

    print("✅ Successfully pushed all JSON files to GitHub.")

except Exception as e:
    print("❌ GitHub push failed:", e)
    error_info = {
        "status": "error",
        "message": "GitHub push failed",
        "error_detail": str(e),
        "traceback": traceback.format_exc()
    }
    with open("status.json", "w", encoding="utf-8") as f:
        json.dump(error_info, f, ensure_ascii=False, indent=4)


# import pandas as pd
# from datetime import datetime, timedelta, timezone
# import json
# import io
# import requests

# # Previous Google Sheet (commented out)
# # SHEET_ID = "1PttAjbFtfvymn9pBpvZdMwNylzQNFVC9yGzvsMl2Vr8"
# # GID = "0"

# # New Google Sheet
# SHEET_ID = "1cmDXt7UTIKBVXBHhtZ0E4qMnJrRoexl2GmDFfTBl0Z4"
# GID = "1882612924"

# csv_url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&gid={GID}"

# try:
#     r = requests.get(csv_url)
#     r.raise_for_status()

#     df = pd.read_csv(io.StringIO(r.text))
#     print("✅ Loaded Google Sheet data:\n", df)

#     data = df.to_dict(orient="records")

#     utc_time = datetime.now(timezone.utc)
#     local_time = utc_time + timedelta(hours=5)  # Pakistan Standard Time (UTC+5)

#     result = {
#         "last_updated": utc_time.strftime("%Y-%m-%d %H:%M:%S UTC"),
#         "local_time": local_time.strftime("%Y-%m-%d %I:%M:%S %p PKT"),
#         "data": data
#     }

#     with open("sheet_backup.json", "w", encoding="utf-8") as f:
#         json.dump(result, f, ensure_ascii=False, indent=4)

#     print("💾 Saved as sheet_backup.json with local time included")

# except Exception as e:
#     print("❌ Error:", e)

