from __future__ import print_function
import datetime
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import csv
from dateutil import parser
from flask import Flask, render_template, request, send_file
import sys

# If modifying these scopes, delete the file token.json.
SCOPES = [
    'https://www.googleapis.com/auth/calendar.readonly'
    ]

# トップ画面を読み込み
app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/privacypolicy')
def privacypolicy():
    return render_template('/privacypolicy.html')

# datetimeを小数に変換
def timedelta_to_decimal(timedelta_obj):
    total_seconds = timedelta_obj.total_seconds()
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60
    decimal_time = hours + (minutes / 60) + (seconds / 3600)
    return round(decimal_time, 2)


# Google APIに接続
@app.route("/export", methods=["post"])
def export():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
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

    service = build('calendar', 'v3', credentials=creds)

    # 入力したデータを格納
    datetime_start = request.form.get('datetime-start')
    datetime_end = request.form.get('datetime-end')
    datetime_start = datetime.datetime.strptime(datetime_start, '%Y-%m-%d')
    datetime_end = datetime.datetime.strptime(datetime_end, '%Y-%m-%d')
    datetime_end = datetime_end + datetime.timedelta(days=1)
    time_min = datetime_start.isoformat() + 'Z'
    time_max = datetime_end.isoformat() + 'Z'

    events_result = service.events().list(
        calendarId='primary',
        timeMin=time_min,
        timeMax=time_max,
        singleEvents=True,
        orderBy='startTime'
    ).execute()
    events = events_result.get('items', [])

    # if not events:
    #     print('No upcoming events found.')
    # for event in events:
    #     start = event['start'].get('dateTime', event['start'].get('date'))
    #     end = event['end'].get('dateTime', event['end'].get('date'))
    #     # start_datetime = parser.parse(start)
    #     start = start.rsplit('+', 1)[0]  # タイムゾーン情報を削除
    #     start_datetime = datetime.datetime.strptime(start, '%Y-%m-%dT%H:%M:%S')
    #     end = end.rsplit('+', 1)[0]  # タイムゾーン情報を削除
    #     end_datetime = datetime.datetime.strptime(end, '%Y-%m-%dT%H:%M:%S')
    #     time = end_datetime - start_datetime
    #     print(start_datetime, end, time, event['summary'])


# CSVを吐き出す処理
    def write_events_to_csv(events, file_path):
        with open(file_path, 'w', newline='') as csvfile:
            fieldnames = ['予定開始日時', '予定実施時間', '予定タイトル']  # 列名
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()  # 列名を書き込む
            
            for event in events:
                start = event['start'].get('dateTime', event['start'].get('date'))
                end = event['end'].get('dateTime', event['end'].get('date'))
                start = start.rsplit('+', 1)[0]  # タイムゾーン情報を削除
                start_datetime = datetime.datetime.strptime(start, '%Y-%m-%dT%H:%M:%S')
                end = end.rsplit('+', 1)[0]  # タイムゾーン情報を削除
                end_datetime = datetime.datetime.strptime(end, '%Y-%m-%dT%H:%M:%S')
                time = end_datetime - start_datetime
                time = timedelta_to_decimal(time)
                writer.writerow({'予定開始日時': start_datetime, '予定実施時間': time, '予定タイトル': event['summary']})
        return file_path
    
    file_path = 'events.csv'
    write_events_to_csv(events, file_path)  # イベントデータをevents.csvに書き込む
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)