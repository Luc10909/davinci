import requests
import json
import datetime
import pytz

# Config
DAVINCI_URL = "https://sp.bs-technik-rostock.de:9090"
USERNAME = "FG51"
PASSWORD = "BS-Technik53510"

def fetch_and_generate_ics():
    url = f"{DAVINCI_URL}/daVinciIS.dll?content=json&username={USERNAME}&password={PASSWORD}"
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        data = response.json()
    except Exception as e:
        print(f"Failed to fetch data: {str(e)}")
        return

    lesson_times = data.get("result", {}).get("displaySchedule", {}).get("lessonTimes", [])
    
    # Setup standard iCalendar header
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//LuciH//DaVinci to GitHub ICS//DE",
        "CALSCALE:GREGORIAN",
        "X-WR-CALNAME:DaVinci Stundenplan",
        "X-WR-TIMEZONE:Europe/Berlin"
    ]
    
    now_dt = datetime.datetime.now(pytz.utc).strftime('%Y%m%dT%H%M%SZ')
    tz = pytz.timezone("Europe/Berlin")

    # Only include events between last Monday and next Sunday
    now = datetime.datetime.now(tz)
    current_weekday = now.weekday()
    start_window = now - datetime.timedelta(days=current_weekday)
    start_window = start_window.replace(hour=0, minute=0, second=0, microsecond=0)
    end_window = now + datetime.timedelta(days=(13 - current_weekday))
    end_window = end_window.replace(hour=23, minute=59, second=59, microsecond=999999)

    for item in lesson_times:
        subject = item.get("subjectCode") or item.get("courseTitle", "Unbekannt")
        start_t = item.get("startTime", "0000")
        end_t = item.get("endTime", "0000")
        dates = item.get("dates", [])
        
        start_h, start_m = int(start_t[:2]), int(start_t[2:4])
        end_h, end_m = int(end_t[:2]), int(end_t[2:4])

        # USER REQUIREMENT: Drop any class starting at or after 14:45
        if start_h * 100 + start_m >= 1445:
            continue
            
        room = item.get("roomCodes", [""])[0] if item.get("roomCodes") else ""
        teacher = ", ".join(item.get("teacherCodes", []))
        
        changes = item.get("changes")
        is_cancelled = False
        is_move = False
        notes = []
        
        if changes:
             caption = changes.get("caption", "")
             ctype = changes.get("type", "")
             
             if changes.get("cancelled") in ["classFree", "movedAway"] or ctype == "cancellation":
                 is_cancelled = True
                 
             if changes.get("cancelled") == "movedAway" or "verschoben" in caption.lower():
                 is_move = True
                 
             if changes.get("modified") == "true" or ctype == "substitution" or "vertretung" in caption.lower():
                 notes.append("VERTRETUNG")
                 
             if caption:
                 notes.append(caption)

        for date_str in dates:
             year = int(date_str[:4])
             month = int(date_str[4:6])
             day = int(date_str[6:8])
             
             dt_start = tz.localize(datetime.datetime(year, month, day, start_h, start_m))
             dt_end = tz.localize(datetime.datetime(year, month, day, end_h, end_m))
             
             if dt_start < start_window or dt_start > end_window:
                 continue
                 
             lines.append("BEGIN:VEVENT")
             
             uid = f"davinci-{subject.replace(' ', '')}-{date_str}-{start_t}@sync"
             lines.append(f"UID:{uid}")
             lines.append(f"DTSTAMP:{now_dt}")
             
             start_utc = dt_start.astimezone(pytz.utc).strftime('%Y%m%dT%H%M%SZ')
             end_utc = dt_end.astimezone(pytz.utc).strftime('%Y%m%dT%H%M%SZ')
             lines.append(f"DTSTART:{start_utc}")
             lines.append(f"DTEND:{end_utc}")
             
             # Formatting title
             prefix = ""
             if is_cancelled and not is_move:
                 prefix = "ENTFÄLLT: "
             elif is_move:
                 prefix = "VERSCHOBEN: "
             elif "VERTRETUNG" in notes:
                 prefix = "VERTRETUNG: "
                 
             lines.append(f"SUMMARY:{prefix}{subject}")
             clean_room = room.replace(',', '\\,')
             lines.append(f"LOCATION:{clean_room}")
             
             desc = f"Lehrer: {teacher}"
             if notes:
                 desc += f"\\nInfo: {', '.join(notes)}"
             lines.append(f"DESCRIPTION:{desc}")
             
             if is_cancelled:
                 lines.append("STATUS:CANCELLED")
             else:
                 lines.append("STATUS:CONFIRMED")
                 
             lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")
    
    with open("schedule.ics", "w", encoding="utf-8") as f:
        f.write("\\r\\n".join(lines))
    print("Successfully generated schedule.ics")

if __name__ == "__main__":
    fetch_and_generate_ics()
