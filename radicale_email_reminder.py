# Copyright (c) 2019 Nader G. Zeid
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/gpl.html>.

import sys
import datetime
import os
import gc
import re
import html
import uuid
import quopri
from smtplib import SMTP
from email.utils import parseaddr
from email.utils import formataddr
from icalendar import Calendar
from icalendar import prop
import dateutil.rrule
from email.header import Header

def initialize_smtp_object(mail_server_host, mail_server_port):
  smtp_object = SMTP(host=mail_server_host, port=mail_server_port, timeout=10)
  smtp_object.ehlo_or_helo_if_needed()
  if(smtp_object.has_extn("starttls")):
    smtp_object.starttls()
  return smtp_object

def get_email_content(vobject):
  email_content = {
    "summary": (vobject["SUMMARY"] if vobject.has_key("SUMMARY") and isinstance(vobject["SUMMARY"], str) else ""),
    "location": (vobject["LOCATION"] if vobject.has_key("LOCATION") and isinstance(vobject["LOCATION"], str) else ""),
    "description": (vobject["DESCRIPTION"] if vobject.has_key("DESCRIPTION") and isinstance(vobject["DESCRIPTION"], str) else "")
  }
  email_content["subject"] = email_content["summary"]
  return email_content

def extract_email_addresses(email_addresses, regex_list):
  output = []
  email_addresses = regex_list["first_email_split"].search(email_addresses)
  if(email_addresses):
    email_addresses = regex_list["whitespace_trim"].sub("", email_addresses.group(2))
    email_addresses = regex_list["second_email_split"].split(email_addresses)
    for email_address in email_addresses:
      email_address = parseaddr(email_address)
      if(len(email_address[1])):
        output.append(formataddr(email_address))
  return output

def trim_email_content(email_content, regex_list):
  email_content["summary"] = regex_list["whitespace_trim"].sub("", email_content["summary"])
  email_content["location"] = regex_list["whitespace_trim"].sub("", email_content["location"])
  email_content["description"] = regex_list["first_email_split"].sub("", email_content["description"])
  email_content["description"] = regex_list["whitespace_trim"].sub("", email_content["description"])
  email_content["subject"] = regex_list["whitespace_clobber"].sub(" ", email_content["subject"])
  email_content["subject"] = email_content["subject"][:70] + (email_content["subject"][70:] and "...")
  return

def send_email(email_addresses, email_content, smtp_object, from_address, report):
  from_field = "From: " + from_address + "\r\n"
  subject = Header(email_content["subject"] + " on " + email_content["start_time"], "utf-8").encode()
  the_rest = "Subject: " + subject + "\r\n"
  the_rest += "MIME-Version: 1.0\r\n"
  boundary = uuid.uuid4().hex
  the_rest += "Content-type: multipart/alternative; boundary=" + boundary + "\r\n\r\n"
  the_rest += "\r\n--" + boundary + "\r\n"
  the_rest += "Content-Type: text/plain; charset=utf-8\r\n"
  the_rest += "Content-Transfer-Encoding: quoted-printable\r\n\r\n"
  a_body = "Summary:\n\n" + email_content["summary"] + "\n\n"
  a_body += "Time:\n\n" + email_content["start_time"] + "\n\n"
  if(len(email_content["location"])):
    a_body += "Location:\n\n" + email_content["location"] + "\n\n"
  if(len(email_content["description"])):
    a_body += "Description:\n\n" + email_content["description"] + "\n\n"
  the_rest += quopri.encodestring(a_body.encode("utf-8")).decode("utf-8")
  the_rest += "\r\n--" + boundary + "\r\n"
  the_rest += "Content-Type: text/html; charset=utf-8\r\n"
  the_rest += "Content-Transfer-Encoding: quoted-printable\r\n\r\n"
  a_body = "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Strict//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd\">\n"
  a_body += "<html xmlns=\"http://www.w3.org/1999/xhtml\">\n"
  a_body += "<head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" /><title>" + html.escape(email_content["subject"]) + "</title></head>\n"
  a_body += "<body>\n"
  a_body += "<div>Summary:</div><ul><li style=\"white-space: pre;\">" + html.escape(email_content["summary"]) + "</li></ul>\n"
  a_body += "<div>Time:</div><ul><li style=\"white-space: pre;\">" + html.escape(email_content["start_time"]) + "</li></ul>\n"
  if(len(email_content["location"])):
    a_body += "<div>Location:</div><ul><li style=\"white-space: pre;\">" + html.escape(email_content["location"]) + "</li></ul>\n"
  if(len(email_content["description"])):
    a_body += "<div>Description:</div><ul><li style=\"white-space: pre;\">" + html.escape(email_content["description"]) + "</li></ul>\n"
  a_body += "</body>\n"
  a_body += "</html>\n"
  the_rest += quopri.encodestring(a_body.encode("utf-8")).decode("utf-8")
  the_rest += "\r\n--" + boundary + "--\r\n"
  for to_address in email_addresses:
    message = from_field
    message += "To: " + to_address + "\r\n"
    message += the_rest
    try:
      smtp_object.sendmail(from_address, to_address, message)
    except:
      report["emails_failed"] += 1
      continue
    report["emails_sent"] += 1
  return

def process_alarm_object(alarm_object, stamp_object, start_object, rrule_object, email_addresses, email_content, email_start_time, minutes_ahead, smtp_object, from_address, report):
  # print(alarm_object.to_ical().decode("utf-8"))
  start_time = None
  if(isinstance(start_object, prop.vDDDTypes)):
    if(isinstance(start_object.dt, datetime.datetime)):
      start_time = start_object.dt
    else:
      start_time = datetime.datetime(year=start_object.dt.year, month=start_object.dt.month, day=start_object.dt.day, tzinfo=stamp_object.dt.tzinfo)
  else:
    start_time = stamp_object.dt
  if(start_time.tzinfo is None):
    start_time = start_time.replace(tzinfo=stamp_object.dt.tzinfo)
  if(isinstance(rrule_object, prop.vRecur)):
    rrule = dateutil.rrule.rrulestr(rrule_object.to_ical().decode('utf-8'), dtstart=start_time)
    start_time = rrule.after(email_start_time)
    if(not isinstance(start_time, datetime.datetime)):
      report["alarms_expired"] += 1
      return
  email_content["start_time"] = start_time.strftime("%a %b %d, %Y %I:%M%p %Z")
  email_content["start_time"] = email_content["start_time"].replace(" 0", " ")
  alarm_time = alarm_object["TRIGGER"].dt
  if(isinstance(alarm_time, datetime.datetime)):
    if(alarm_time.tzinfo is None):
      alarm_time = alarm_time.replace(tzinfo=start_time.tzinfo)
  else:
    alarm_time = start_time + alarm_time;
  email_end_time = email_start_time + minutes_ahead
  # print([alarm_time.isoformat(), email_start_time.isoformat(), email_end_time.isoformat()])
  if(alarm_time >= email_start_time):
    if(alarm_time < email_end_time):
      send_email(email_addresses, email_content, smtp_object, from_address, report)
      report["alarms_triggered"] += 1
    else:
      report["alarms_pending"] += 1
  else:
    report["alarms_expired"] += 1
  return

def process_calendar_object(calendar_object, email_start_time, minutes_ahead, smtp_object, from_address, regex_list, report):
  # print(calendar_object.to_ical().decode("utf-8"))
  for event_object in calendar_object.walk('vevent'):
    # print(event_object.to_ical().decode("utf-8"))
    stamp_object = event_object["DTSTAMP"]
    start_object = event_object["DTSTART"] if event_object.has_key("DTSTART") else None
    rrule_object = event_object["RRULE"] if event_object.has_key("RRULE") else None
    email_content = get_email_content(event_object)
    email_addresses = []
    if(len(email_content["description"])):
      email_addresses = extract_email_addresses(email_content["description"], regex_list)
    trim_email_content(email_content, regex_list)
    for alarm_object in event_object.walk('valarm'):
      report["event_alarms"] += 1
      process_alarm_object(alarm_object, stamp_object, start_object, rrule_object, email_addresses, email_content, email_start_time, minutes_ahead, smtp_object, from_address, report)
  for todo_object in calendar_object.walk('vtodo'):
    # print(todo_object.to_ical().decode("utf-8"))
    stamp_object = todo_object["DTSTAMP"]
    email_content = get_email_content(todo_object)
    email_addresses = []
    if(len(email_content["description"])):
      email_addresses = extract_email_addresses(email_content["description"], regex_list)
    trim_email_content(email_content, regex_list)
    for alarm_object in todo_object.walk('valarm'):
      report["todo_alarms"] += 1
      process_alarm_object(alarm_object, stamp_object, None, None, email_addresses, email_content, email_start_time, minutes_ahead, smtp_object, from_address, report)
  return

def process_calendar_file(calendar_file, email_start_time, minutes_ahead, smtp_object, from_address, regex_list, report):
  try:
    file_handle = open(calendar_file, "rb")
  except:
    report["calendar_file_access_errors"] += 1
    return
  try:
    calendar_object = file_handle.read()
    calendar_object = calendar_object.decode("utf-8")
    calendar_object = Calendar.from_ical(calendar_object)
  except:
    file_handle.close()
    report["calendar_file_format_errors"] += 1
    return
  file_handle.close()
  process_calendar_object(calendar_object, email_start_time, minutes_ahead, smtp_object, from_address, regex_list, report)
  return

def process_calendar_directory(calendar_directory, minutes_ahead, mail_server_host, mail_server_port, from_address, report):
  if(not os.path.isdir(calendar_directory)):
    print(os.path.basename(__file__) + ": the path \"" + calendar_directory + "\" is invalid!", file=sys.stderr)
    exit(1)
  minutes_ahead = int(minutes_ahead)
  if(not (minutes_ahead > 0)):
    print(os.path.basename(__file__) + ": the minutes ahead \"" + str(minutes_ahead) + "\" must be positive!", file=sys.stderr)
    exit(1)
  minutes_ahead = datetime.timedelta(minutes=minutes_ahead)
  email_start_time = datetime.datetime.now(datetime.timezone.utc)
  email_start_time = datetime.datetime(
    year=email_start_time.year,
    month=email_start_time.month,
    day=email_start_time.day,
    hour=email_start_time.hour,
    minute=email_start_time.minute,
    second=0,
    microsecond=0,
    tzinfo=email_start_time.tzinfo
  )
  regex_list = {
    "whitespace_trim": re.compile("^\\s+|\\s+$", re.S|re.U),
    "whitespace_clobber": re.compile("\\s+", re.S|re.U),
    "first_email_split": re.compile("^[^\\S\\r\\n]*[Nn][Oo][Tt][Ii][Ff][Yy]:[^\\S\\r\\n]*(\\r\\n|\\r|\\n)(.+?)(\\r\\n|\\r|\\n)-", re.S|re.U),
    "second_email_split": re.compile("\\s*\\r\\s*|\\s*\\n\\s*", re.S|re.U),
  }
  smtp_object = initialize_smtp_object(mail_server_host, mail_server_port)
  for root, dirs, files in os.walk(calendar_directory):
    for one_file in files:
      if(one_file.endswith(".ics")):
        report["calendar_files"] += 1
        process_calendar_file(root + "/" + one_file, email_start_time, minutes_ahead, smtp_object, from_address, regex_list, report)
        gc.collect()
  return

if(len(sys.argv) != 6):
  print(os.path.basename(__file__) + ": Missing operands!\n" + os.path.basename(__file__) + " <directory> <minutes ahead> <mail server host> <mail server port> <from address>", file=sys.stderr)
  exit(1)
report = {
  "calendar_files": 0,
  "calendar_file_access_errors": 0,
  "calendar_file_format_errors": 0,
  "event_alarms": 0,
  "todo_alarms": 0,
  "alarms_expired": 0,
  "alarms_triggered": 0,
  "alarms_pending": 0,
  "emails_sent": 0,
  "emails_failed": 0
}
process_calendar_directory(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5], report)
print(report)
