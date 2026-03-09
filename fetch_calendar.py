import requests
import json
import datetime
import pytz
import re

# Config
DAVINCI_URL = "https://sp.bs-technik-rostock.de:9090"
USERNAME = "FG51"
PASSWORD = "BS-Technik53510"

# ==========================================
# FILTER CONFIGURATION (Sync with Dashboard)
# ==========================================
# Diese Fächer werden komplett ausgeblendet.
IGNORED_SUBJECTS = ["mt", "mtl", "ku", "kunst", "fra", "frf", "ruf"]

# Bevorzugte Fächer bei Parallelkursen (Wahlpflicht)
PREFERRED_SUBJECTS = ["rua", "russisch"]
# ==========================================

def normalize_subject(subj):
    """Remove +, (z), (D) etc to find the core subject for merging"""
    s = subj.replace("+", "")
    s = re.sub(r'\(.*?\)', '', s)
    return s.strip().lower()

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
    
    tz = pytz.timezone("Europe/Berlin")
    now = datetime.datetime.now(tz)
    
    # 1. Collect and filter all events
    events_by_timeslot = {}
    
    current_weekday = now.weekday()
    start_window = (now - datetime.timedelta(days=current_weekday)).replace(hour=0, minute=0, second=0, microsecond=0)
    end_window = (now + datetime.timedelta(days=13 - current_weekday)).replace(hour=23, minute=59, second=59, microsecond=999999)

    for item in lesson_times:
        subject = item.get("subjectCode") or item.get("courseTitle", "Unbekannt")
        
        if subject.lower() in IGNORED_SUBJECTS:
            continue
            
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
                 
             if (changes.get("modified") == "true" or ctype == "substitution" or 
                 "vertretung" in caption.lower()):
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
             
             # Prep base event object
             event_obj = {
                 "title": subject,
                 "raw_subject": subject,
                 "start": dt_start,
                 "end": dt_end,
                 "location": room,
                 "teacher": teacher,
                 "is_cancelled": is_cancelled,
                 "is_move": is_move,
                 "notes": notes.copy(),
                 "date_str": date_str,
                 "start_time_str": start_t
             }
             
             # Group by exact time interval to handle parallel blocks
             group_key = f"{date_str}_{start_t}_{end_t}"
             if group_key not in events_by_timeslot:
                 events_by_timeslot[group_key] = []
             events_by_timeslot[group_key].append(event_obj)

    # 2. Resolve parallel blocks inside identical timeslots (Wahlpflicht)
    resolved_events = []
    for key, slot_events in events_by_timeslot.items():
        if len(slot_events) == 1:
            resolved_events.append(slot_events[0])
            continue
            
        chosen_event = None
        for evt in slot_events:
            if any(pref in evt["raw_subject"].lower() for pref in PREFERRED_SUBJECTS):
                chosen_event = evt
                break
                
        if not chosen_event:
            non_cancelled = [e for e in slot_events if not e["is_cancelled"]]
            chosen_event = non_cancelled[0] if non_cancelled else slot_events[0]
                
        resolved_events.append(chosen_event)

    # 3. Sort by time and merge consecutive slots of the SAME subject
    day_events = {}
    for evt in resolved_events:
        day_date = evt["date_str"]
        if day_date not in day_events:
            day_events[day_date] = []
        day_events[day_date].append(evt)
        
    merged_events = []
    for day_date, events_in_day in day_events.items():
        events_in_day.sort(key=lambda x: x["start"])
        
        current_block = None
        for evt in events_in_day:
            if not current_block:
                current_block = evt
            else:
                norm_curr = normalize_subject(current_block["raw_subject"])
                norm_next = normalize_subject(evt["raw_subject"])
                
                if norm_curr == norm_next and current_block["end"] == evt["start"]:
                    # Merge!
                    current_block["end"] = evt["end"]
                    # If one part is cancelled, but not the other, or both etc.
                    if evt["is_cancelled"] != current_block["is_cancelled"]:
                         current_block["notes"].append("Teilweise Ausfall")
                    
                    # Merge notes
                    for n in evt["notes"]:
                        if n not in current_block["notes"]:
                            current_block["notes"].append(n)
                    continue
                
                merged_events.append(current_block)
                current_block = evt
        if current_block:
            merged_events.append(current_block)

    # 4. Generate ICS lines
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//LuciH//DaVinci to GitHub ICS//DE",
        "CALSCALE:GREGORIAN",
        "X-WR-CALNAME:DaVinci Stundenplan",
        "X-WR-TIMEZONE:Europe/Berlin"
    ]
    
    now_utc_str = datetime.datetime.now(pytz.utc).strftime('%Y%m%dT%H%M%SZ')

    for evt in merged_events:
        lines.append("BEGIN:VEVENT")
        
        # Use first start time or something stable for UID
        uid = f"davinci-{normalize_subject(evt['raw_subject'])}-{evt['date_str']}-{evt['start'].strftime('%H%M')}@sync"
        lines.append(f"UID:{uid}")
        lines.append(f"DTSTAMP:{now_utc_str}")
        
        lines.append(f"DTSTART:{evt['start'].astimezone(pytz.utc).strftime('%Y%m%dT%H%M%SZ')}")
        lines.append(f"DTEND:{evt['end'].astimezone(pytz.utc).strftime('%Y%m%dT%H%M%SZ')}")
        
        # Formatting title
        prefix = ""
        if evt["is_cancelled"] and not evt["is_move"]:
            prefix = "ENTFÄLLT: "
        elif evt["is_move"]:
            prefix = "VERSCHOBEN: "
        elif "VERTRETUNG" in evt["notes"]:
            prefix = "VERTRETUNG: "
            
        lines.append(f"SUMMARY:{prefix}{evt['title']}")
        clean_room = evt["location"].replace(',', '\\,')
        lines.append(f"LOCATION:{clean_room}")
        
        desc = f"Lehrer: {evt['teacher']}"
        if evt["notes"]:
            desc += f"\\nInfo: {', '.join(evt['notes'])}"
        lines.append(f"DESCRIPTION:{desc}")
        
        if evt["is_cancelled"]:
            lines.append("STATUS:CANCELLED")
        else:
            lines.append("STATUS:CONFIRMED")
            
        lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")
    
    with open("schedule.ics", "w", encoding="utf-8") as f:
        f.write("\r\n".join(lines))
    print(f"Successfully generated schedule.ics with {len(merged_events)} blocks.")

if __name__ == "__main__":
    fetch_and_generate_ics()
