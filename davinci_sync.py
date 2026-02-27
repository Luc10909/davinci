import requests
import json
import urllib3
import logging
from datetime import datetime, timedelta
import pytz

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_davinci_schedule(base_url, username, password):
    url = f"{base_url}/daVinciIS.dll"
    if base_url.endswith("/"):
        url = f"{base_url}daVinciIS.dll"
    params = {
        "content": "json",
        "username": username,
        "password": password
    }
    
    try:
        logging.info(f"Fetching schedule from {url}...")
        response = requests.get(url, params=params, verify=False, timeout=15)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        logging.error(f"Failed to fetch data: {e}")
        return None

def generate_ics(schedule_data, output_file="davinci_schedule.ics"):
    if not schedule_data or "result" not in schedule_data or "displaySchedule" not in schedule_data["result"]:
        logging.error("Invalid schedule data format.")
        return False

    display_schedule = schedule_data["result"]["displaySchedule"]
    lesson_times = display_schedule.get("lessonTimes", [])
    
    # German Timezone (relevant for the schedule)
    tz = pytz.timezone("Europe/Berlin")
    
    ics_lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//DaVinci to ICS//NONSGML v1.0//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        "X-WR-CALNAME:DaVinci Stundenplan",
        "X-WR-TIMEZONE:Europe/Berlin"
    ]
    
    for lesson in lesson_times:
        subject = lesson.get("subjectCode", lesson.get("courseTitle", "Unbekannt"))
        start_time_str = lesson.get("startTime", "0000") # HHmm
        end_time_str = lesson.get("endTime", "0000") # HHmm
        
        rooms = lesson.get("roomCodes", [])
        room = rooms[0] if rooms else ""
        
        teachers = lesson.get("teacherCodes", [])
        teacher = ", ".join(teachers)

        # Changes interpretation
        note = ""
        is_cancelled = False
        is_substitution = False
        changes = lesson.get("changes", None)
        
        if changes:
            caption = changes.get("caption", "")
            change_type = changes.get("type", "")
            if caption:
                note += f"\nHinweis: {caption}"
            
            if changes.get("cancelled") == "classFree" or change_type == "cancellation":
                is_cancelled = True
            if changes.get("modified") == "true" or change_type == "substitution" or "vertretung" in caption.lower():
                is_substitution = True

        title_prefix = ""
        if is_cancelled:
            title_prefix = "ENTFÄLLT: "
        elif is_substitution:
            title_prefix = "VERTRETUNG: "

        final_title = f"{title_prefix}{subject}"
        description = f"Lehrer: {teacher}{note}"

        dates = lesson.get("dates", [])
        for date_str in dates:
             # date_str: YYYYMMDD
             start_dt_str = f"{date_str}{start_time_str}"
             end_dt_str = f"{date_str}{end_time_str}"
             
             try:
                 start_dt = datetime.strptime(start_dt_str, "%Y%m%d%H%M")
                 end_dt = datetime.strptime(end_dt_str, "%Y%m%d%H%M")
                 
                 # Localize
                 start_dt = tz.localize(start_dt)
                 end_dt = tz.localize(end_dt)
                 
                 # To UTC for ICS
                 start_utc = start_dt.astimezone(pytz.utc).strftime("%Y%m%dT%H%M%SZ")
                 end_utc = end_dt.astimezone(pytz.utc).strftime("%Y%m%dT%H%M%SZ")
                 
                 # Create a unique ID for this specific lesson and date so calendar apps can OVERWRITE it instead of duplicating.
                 # E.g. uid = subjectCode-date-startTime@davinci.sync
                 uid = f"davinci-{subject.replace(' ', '')}-{date_str}-{start_time_str}@sync"
                 
                 stamp = datetime.now(pytz.utc).strftime("%Y%m%dT%H%M%SZ")

                 desc_escaped = description.replace('\n', '\\n')
                 ics_lines.extend([
                     "BEGIN:VEVENT",
                     f"UID:{uid}",
                     f"DTSTAMP:{stamp}",
                     f"DTSTART:{start_utc}",
                     f"DTEND:{end_utc}",
                     f"SUMMARY:{final_title}",
                     f"LOCATION:{room}",
                     f"DESCRIPTION:{desc_escaped}",
                 ])
                 
                 # If cancelled, modern calendars can use STATUS:CANCELLED
                 if is_cancelled:
                     ics_lines.append("STATUS:CANCELLED")
                     
                 ics_lines.append("END:VEVENT")

             except Exception as e:
                 logging.warning(f"Error parsing date {date_str}: {e}")
                 
    ics_lines.append("END:VCALENDAR")
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write('\n'.join(ics_lines))
        
    logging.info(f"Successfully wrote {output_file}")
    return True

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Fetch DaVinci Schedule and convert to ICS")
    parser.add_argument("--url", default="https://sp.bs-technik-rostock.de:9090", help="Base URL of DaVinci server")
    parser.add_argument("--user", default="FG51", help="Username")
    parser.add_argument("--password", default="BS-Technik53510", help="Password")
    parser.add_argument("--output", default="davinci_schedule.ics", help="Output ICS file name")
    
    args = parser.parse_args()
    
    data = get_davinci_schedule(args.url, args.user, args.password)
    if data:
        generate_ics(data, args.output)
