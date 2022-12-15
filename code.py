# Importing necessary libraries and files
from __future__ import print_function

import datetime
import os.path
import iso8601
import matplotlib
import matplotlib.pyplot as plt
import pandas as pd
import math
import xlsxwriter
import csv


from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


matplotlib.use('TkAgg')

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']


def main():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('calendar', 'v3', credentials=creds)

        # Create a calendar list and print their names
        page_token = None
        print('Getting the existing calendars')
        while True:
            calendar_list = service.calendarList().list(pageToken=page_token).execute()
            for calendar_list_entry in calendar_list['items']:
                print(calendar_list_entry['summary'])
            page_token = calendar_list.get('nextPageToken')
            if not page_token:
                break

        # Set up the analyzed time frame
        start_of_week = "2022-10-10T00:00:00-07:00"
        end_of_week = "2022-10-16T12:00:00-07:00"

        # Get the sum for each calendar
        dict_of_time = {}

        for calendar_list_entry in calendar_list['items']:
            if calendar_list_entry['id'] == '---': # Making an exception (if necessary)
                continue
            else: # get events for the week
                events = service.events().list(calendarId='{}'.format(calendar_list_entry['id']),
                                               timeMin=start_of_week,
                                               timeMax=end_of_week, pageToken=page_token).execute()
                sum_hours = 0
                sum_minutes = 0
                print(calendar_list_entry['summary'])
                for event in events['items']:
                    start = iso8601.parse_date(event['start'].get('dateTime'))
                    end = iso8601.parse_date(event['end'].get('dateTime'))
                    if end.hour == 0:
                        total_hours = abs(24 - start.hour)
                    else:
                        total_hours = abs(end.hour - start.hour)
                    total_minutes = abs(end.minute - start.minute)
                    sum_hours += total_hours
                    sum_minutes += total_minutes
                    print(f"{event['summary']} took {total_hours} hours and {total_minutes} minutes")
                    # calculate total hours
                    if sum_minutes >= 60:
                        sum_hours += 1
                        sum_minutes -= 60

                # Create a dictionary from calendars and their total time
                dict_of_time[calendar_list_entry['summary']] = sum_hours + sum_minutes / 60
                print(
                    f'You spend the total of {sum_hours} hours and {sum_minutes} minutes on {calendar_list_entry["summary"]} this week')

        headers = ['Calendar', 'Time']

        # Create a csv file from the dictionary - with total time for each calendar
        with open('calendar.csv', 'w') as f:
            f.write('{0},{1}\n'.format(headers[0], headers[1]))
            [f.write('{0},{1}\n'.format(key, value)) for key, value in dict_of_time.items()]

        # Make a list from the csv file (to later put into the Excel report)
        arr_table = []
        with open('calendar.csv') as file:
            reader = csv.reader(file)
            for row in reader:
                arr_table.append(row)

        # Separate the total time spent in a week by hours and minutes
        weekly_mins, weekly_hours = math.modf(sum(dict_of_time.values()))
        weekly_mins *= 60

        # Make the pie chart
        df = pd.read_csv('calendar.csv')
        calendar_data = df['Calendar']
        time_data = df['Time']
        colors = ['#ecd5e3', '#eceae4', '#97c1a9', '#cbbacb', '#abdee6', '#f3b0c3']

        # Format the labels in the pie chart
        def autopct_format(values):
            def my_format(pct):
                total = sum(values)
                val = int(round(pct * total / 100.0))
                return '~{v:d} hours\n({:.1f}%)'.format(pct, v=val)

            return my_format

        plt.pie(time_data, labels=calendar_data, colors=colors, autopct=autopct_format(time_data),
                textprops={'fontsize': 8})
        plt.title(
            f"You've spent the total of {int(weekly_hours)} hours and {int(weekly_mins)} minutes being productive")
        plt.savefig('weekly_report.png')

        # Create a Report in an Excel file
        workbook = xlsxwriter.Workbook('report2.xlsx')
        worksheet = workbook.add_worksheet('main sheet')
        bold = workbook.add_format({'bold': True})
        worksheet.write('A1', 'Productivity Report for {}/{}-{}/{}'.format(iso8601.parse_date(start_of_week).day,
                                                                           iso8601.parse_date(start_of_week).month,
                                                                           iso8601.parse_date(end_of_week).day,
                                                                           iso8601.parse_date(end_of_week).month), bold)
        worksheet.set_column('A:A', 20)
        row = 1

        # Put the table there
        for i in arr_table:
            column = 0
            for j in i:
                worksheet.write(row, column, j)
                column += 1
            row += 1

        # Set and check against goals for the week
        goal = 20
        max_calendar = max(dict_of_time, key=dict_of_time.get)
        worksheet.write('A10', 'You have spent most of your time on {}'.format(max_calendar), bold)
        python_mins, python_hours = math.modf(dict_of_time['Learning Python & ML'])
        python_mins *= 60
        worksheet.write('A12',
                        f'You have spent {int(python_hours)} hours and {int(python_mins)} minutes learning Python',
                        bold)

        if python_hours < goal:
            worksheet.write('A13', f'This is {int(20 - python_hours)} hours less than your goal. You can do better!')
        else:
            worksheet.write('A13', f'This is {abs(int(20 - python_hours))} hours more than your goal. Good job!')

        # Insert the image
        worksheet.insert_image('F1', 'weekly_report.png')

        # Close the file
        workbook.close()

    except HttpError as error:
        print('An error occurred: %s' % error)

if __name__ == '__main__':
    main()


