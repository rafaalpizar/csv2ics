#!/usr/bin/env python3

import csv
import os
import sys
from datetime import timedelta
from pathlib import Path
from uuid import uuid4
import dateparser
import re

from ical.calendar import Calendar
from ical.calendar_stream import IcsCalendarStream
from ical.event import Event
from ical.types import Recur
from ical.types.recur import Frequency

if __name__ == "__main__":
    csv_file_input = sys.argv[1]
    base_name = csv_file_input[:-4]

    cal = Calendar()
    with open(csv_file_input, mode="r") as csv_file:
        csv_reader = csv.DictReader(csv_file)
        line_count = 0
        for row in csv_reader:
            if line_count >= 0:
                start, end = (
                    dateparser.parse(row["Start Date"]+" "+row["Start Time"]),
                    dateparser.parse(row["End Date"]+" "+row["End Time"]),
                )
                if not end:  # bad/no end date, make it tomorrow
                    end = start + timedelta(days=1)
                else:
                    end = end
                print(f"{start} ➡️ {end}: {row['Description'], row['Location']}")
                until = dateparser.parse(row["Until Date"]+" "+row["End Time"])
                # TODO: Set frequency from csv file
                rrule = Recur(freq=Frequency.WEEKLY, until=until)
                cal.events.append(
                    Event(
                        start=start,
                        end=end,
                        summary=row["Description"],
                        transparency=row["Transp"],
                        location=row["Location"],
                        rrule=rrule
                    )
                )
            line_count += 1
    print(f"Processed {line_count} lines.")

    ics_string = IcsCalendarStream.calendar_to_ics(cal)

    # Extra fields
    # regex = r"^END:VEVENT"

    # To be added based on a condition for all day and outlook
    # subst = "X-MICROSOFT-CDO-ALLDAYEVENT:TRUE\\nEND:VEVENT"
    # ics_string = re.sub(regex, subst, ics_string, 0, re.MULTILINE)

    filename = Path(f"{base_name}.ics")
    with filename.open("w") as ics_file:
        ics_file.write(ics_string)
