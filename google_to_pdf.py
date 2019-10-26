from __future__ import print_function

import sys
import csv
from datetime import datetime

from dateutil import parser
from googleapiclient.discovery import build
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import *

from credentials import fetch_credentials



calendar_id_minejendom = "5drdnd66pka1v7m5n2o7usn1n0@group.calendar.google.com"
calendar_id_sidecourt = "qmb4l18ht68lehg6mgrdv9o2b8@group.calendar.google.com"
calendar_id_freelance = "fqrqoviab1djt56445nid42pfs@group.calendar.google.com"
calendar_name_list = ["minejendom", "sidecourt", "freelance"]
calendar_id_list = [calendar_id_minejendom, calendar_id_sidecourt, calendar_id_freelance]

def main():
    """ Fetch google calendar events and write timestamps to csv and pdf """

    """ CALENDAR """
    print("Please pick a calender to use")
    for index in range(0,len(calendar_id_list)): 
        print("" + str(index) + ") " + calendar_name_list[index])
    calender_index = int(input())
    print("Using calender")
    print(calendar_name_list[calender_index])
    print("")

    """ START DATE """
    print("Please specify a start date")
    print("year, month, day, hour, minute")
    default_start_date = [2019, 1, 20, 0, 0]
    print("The value defaults to")
    print(str(default_start_date))

    input_date_string = input()
    input_date_string_array = [x.strip() for x in input_date_string.split(',')] if input_date_string else []
    input_date_string_array_len = len(input_date_string_array)
    start_date = [int(input_date_string_array[x]) if input_date_string_array_len > x else default_start_date[x]  for x in range(0, 5)]
    print("Using the following start date")
    print(str(start_date))
    print("")


    """ END DATE """
    print("Please specify an end date")
    print("year, month, day, hour, minute")
    default_end_date = [2019, 2, 19, 23, 59]
    print("The value defaults to")
    print(str(default_end_date))

    end_date_string = input()
    end_date_string_array = [x.strip() for x in end_date_string.split(',')] if end_date_string else []
    end_date_string_array_len = len(end_date_string_array)
    end_date = [int(end_date_string_array[x]) if end_date_string_array_len > x else default_end_date[x]  for x in range(0, 5)]
    print("Using the following end date")
    print(str(end_date))
    print("")


    print("Please a search name for the events")
    print("If empty all events will be used")
    search_string = input()
    print("")


    print("Initialization complete with the following values")
    print("Calender: " + str(calendar_name_list[calender_index]))
    print("From: " + str(start_date))
    print("To: " + str(end_date))
    print("Search: " + str(search_string))
    input("Press any key to confirm and start search")
    print("")


    print('Fetching credentials')
    creds = fetch_credentials()
    service = build('calendar', 'v3', credentials=creds)


    # Define script configs
    name = "Michael Guldborg Consulting"
    calendar_id = calendar_id_list[calender_index] # calendar_id_minejendom # calendar_id_sidecourt # 
    time_min = datetime(start_date[0], start_date[1], start_date[2], start_date[3], start_date[4], 00)
    time_max = datetime(end_date[0], end_date[1], end_date[2], end_date[3], end_date[4], 00)
    time_min_iso = time_min.isoformat() + 'Z'  # 'Z' indicates UTC time
    time_max_iso = time_max.isoformat() + 'Z'  # 'Z' indicates UTC time

    print("Fetching events from calendar:\n{}".format(calendar_id))
    print("From:\t{}\nTo:  \t{}\n".format(time_min_iso, time_max_iso))
    response = service.events().list(calendarId=calendar_id, orderBy='startTime',
                                     timeMin=time_min_iso, timeMax=time_max_iso,
                                     singleEvents=True).execute()

    print("Parsing event response")
    calendar_name = response['summary']
    events = response['items']
    parsed_events = parse_response(events, search=search_string)

    print("Writing data to file")
    filename = "output/{}-{:02d}-{:02d}".format(str(calendar_name).lower(), time_min.month, time_max.month)
    filename = filename + "-{}".format(search_string) if search_string else filename
    headers = [
        name,
        'Fra: {}'.format(time_min.strftime("%d-%m-%Y")),
        'Til: {}'.format(time_max.strftime("%d-%m-%Y")),
    ]

    # csv_filename = filename + ".csv"
    # write_csv(csv_filename, parsed_events)
    # csv_file = open(csv_filename, "r")
    # data = list(csv.reader(csv_file, delimiter=';'))
    pdf_filename = filename + ".pdf"
    write_pdf(pdf_filename, parsed_events, headers=headers)


def parse_response(events, search=""):
    parsed_events = [['Dato', 'Start', 'Slut', 'Varighed']]

    total_duration = None
    for event in events:
        summary = event['summary']
        if search and summary.lower().find(search) is -1:
            print("Skipped {}".format(summary))
            continue


        start = parser.parse(event['start']['dateTime'])
        end = parser.parse(event['end']['dateTime'])

        date = start.date()
        start_time = "{:02d}:{:02d}".format(start.hour, start.minute)
        end_time = "{:02d}:{:02d}".format(end.hour, end.minute)
        duration = end - start

        parsed_events.append([date.strftime("%d-%m-%Y"), start_time, end_time, str(duration)])
        total_duration = total_duration + duration if total_duration else duration
        print("{}, {}, {}, {}, {}".format(summary, date, start_time, end_time, duration))

    # calculate total duration
    days = total_duration.days
    hours, remainder = divmod(total_duration.seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    print("days: {}, hours: {}, minutes: {}".format(days, hours, minutes))

    hours_total = days * 24 + hours + minutes / 60
    parsed_events.append(["", "", "Total", str(hours_total)])
    print("hours_total: {}\n".format(hours_total))
    return parsed_events


def write_csv(filename, data, headers=None):
    print("Writing to file: {}".format(filename))
    file = open(filename, mode='w', newline='')  # to prevent csv writer making double newline
    writer = csv.writer(file, delimiter=';')
    if headers:
        for header in headers:
            writer.writerow([header, "", "", ""])
    writer.writerows(data)
    file.close()


def write_pdf(filename, data, headers=None):
    print("Writing to file: {}".format(filename))
    elements = []

    # PDF Text
    # PDF Text - Styles
    styles = getSampleStyleSheet()
    styleNormal = styles['Normal']

    # PDF Header
    for header in headers:
        elements.append(Paragraph(header, styleNormal))
    if headers: elements.append(Spacer(3 * cm, 0.5 * cm))

    # PDF Table
    # PDF Table - Styles
    # [(start_column, start_row), (end_column, end_row)]
    all_cells = [(0, 0), (-1, -1)]
    header_cells = [(0, 0), (-1, 0)]
    table_style = TableStyle([
        ('VALIGN', all_cells[0], all_cells[1], 'TOP'),
        ('ALIGN', all_cells[0], all_cells[1], 'LEFT'),
        ('LINEBELOW', header_cells[0], header_cells[1], 1, colors.black),
    ])

    # PDF Table - Column Widths
    col_widths = [
        6 * cm,  # Column 0
        2 * cm,  # Column 1
        2 * cm,  # Column 2
        2 * cm,  # Column 3
    ]

    # PDF Table
    for index, row in enumerate(data):
        for col, val in enumerate(row):
            data[index][col] = Paragraph(val, styles['Normal'])

    # Add table to elements
    t = Table(data, colWidths=col_widths, hAlign='LEFT')
    t.setStyle(table_style)
    elements.append(t)

    # Generate PDF
    archivo_pdf = SimpleDocTemplate(
        filename,
        pagesize=letter,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=28)
    archivo_pdf.build(elements)
    print('PDF Generated!')


if __name__ == '__main__':
    main()
