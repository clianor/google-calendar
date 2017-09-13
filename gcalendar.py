import httplib2
import os
import re
from googleapiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import datetime

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Calendar'

time_format = '{d[0]}-{d[1]}-{d[2]}T{d[3]}:{d[4]}:00+09:00'
p = re.compile(r'(\d+)[\.](\d+)[\.](\d+)[-](\d+)[:](\d+)')

def get_credentials():
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'calendar-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:
            credentials = tools.run_flow(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

class GCalendar:
    def __init__(self):
        credentials = get_credentials()
        http = credentials.authorize(httplib2.Http())
        self.service = discovery.build('calendar', 'v3', http=http)
        self.calL = self.service.calendarList().list().execute()['items']
    
    def get_Calendar(self):
        calNL = [l.get('summary') for l in self.calL]
        return calNL

    def insert_Calendar(self, summary='primary', body=None, start=None, end=None):
        calendarInfo = [l for l in self.calL if l.get('summary') == summary]
        if len(calendarInfo) != 1:
            print('실패하였습니다'); return False
        else:
            try:
                m = p.search(start)
                start_data = [m.group(i) for i in range(1, 6)]
                m = p.search(end)
                end_data = [m.group(i) for i in range(1, 6)]

                start = time_format.format(d=start_data)
                end = time_format.format(d=end_data)

                calendarId = calendarInfo[0].get('id')
                body = {
                    "start": {"dateTime": start},
                    "end": {"dateTime": end},
                    "summary": body
                }
                self.service.events().insert(
                    calendarId=calendarId,
                    body=body
                ).execute()

                print('캘린더 추가에 성공하였습니다.')
                return True
            except Exception as e:
                print(e)
                print('실패하였습니다'); return False

    def list_Calendar(self, summary='primary', maxResult=10):
        calendarInfo = [l for l in self.calL if l.get('summary') == summary]

        if len(calendarInfo) != 1:
            print('이벤트 검색에 실패하였습니다.'); return False
        else:
            try:
                now = datetime.datetime.utcnow().isoformat() + 'Z'  # 현재 시간
                calendarId = calendarInfo[0].get('id')
                eventResult = self.service.events().list(
                    calendarId=calendarId,
                    timeMin=now,
                    maxResults=maxResult,
                    singleEvents=True,
                    orderBy='startTime'
                ).execute()

                events = eventResult.get('items', [])
                for event in events:
                    if 'date' in event.get('start', []):
                        print(event['start']['date'] + '(종일)' + '\t' + event['summary'])
                    else:
                        print(event['start'].get('dateTime') + '\t' + event['summary'])

                return True
            except Exception as e:
                print(e)
                print('이벤트 검색에 실패하였습니다.'); return False


if __name__ == '__main__':
    gcalendar = GCalendar()
    #gcalendar.insert_Calendar(summary='연구실-개인 일정', body='김지원이다', start='2017.09.13-06:00', end='2017.09.13-12:00')
    gcalendar.list_Calendar(summary='연구실-개인 일정')