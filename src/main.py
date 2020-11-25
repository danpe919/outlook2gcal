import datetime
from datetime import date

import win32com.client
import requests

from gcal import register

TIMEZONE = "Asia/Tokyo"

def to_datetime(obj):
    time_str = str(obj).split("+")[0]
    return datetime.datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")


def get_calendar(begin: datetime, end: datetime):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    calendar = outlook.getDefaultFolder(9).Items
    calendar.IncludeRecurrences = True
    calendar.Sort("[Start]")

    b_str = begin.strftime("%Y-%m-%d")
    e_str = end.strftime("%Y-%m-%d")
    restriction = f"[Start] >= '{b_str}' AND [END] <= '{e_str}'"
    calendar = calendar.Restrict(restriction)
    return calendar


def get_event_list():
    begin = datetime.datetime.now() + datetime.timedelta(days=1) # tomorrow
    end = begin + datetime.timedelta(days=1)
    cal = get_calendar(begin, end)

    events = []
    for meeting in cal:
        start = to_datetime(meeting.start)
        end = to_datetime(meeting.end)
        print(start, end)
        if start.hour > 12:
            continue # Ignore P.M. events
        events.append(
            {
                # "summary": meeting.subject,
                "summary": "予定あり",
                "start": {"dateTime": start.strftime("%Y-%m-%dT%H:%M:%S.%f+0900"), "timeZone": TIMEZONE},
                "end": {"dateTime": end.strftime("%Y-%m-%dT%H:%M:%S.%f+0900"), "timeZone": TIMEZONE},
            }
        )
    return events


def main():
    event_list = get_event_list()
    for event in event_list:
        register(event)


if __name__ == "__main__":
    main()


def register_google_calendar(body: dict):
    print("Now registering...")
    r = requests.post(
        "https://www.googleapis.com/calendar/v3/calendars/{calendar_id}/events?key={API_KEY}",
        body,
    )
    r.raise_for_status()
    print("Done")
    print(r.json())
