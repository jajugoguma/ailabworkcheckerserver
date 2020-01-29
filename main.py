import telegram
import gspread
from PyQt5.QtCore import QDateTime
from oauth2client.service_account import ServiceAccountCredentials
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.interval import IntervalTrigger
from apscheduler.triggers.cron import CronTrigger
import calendar
import requests
import datetime
import schedule

class Request:
    def __init__(self):
        self.URL = 'url'
        self.OPERATION = 'getHoliDeInfo'  # 국경일 + 공휴일 정보 조회 오퍼레이션
        self.SERVICEKEY = 'key'
        self.solYear = time.YEAR
        self.solMonth = str(time.MONTH).zfill(2)
        self.PARAMS = {'solYear': self.solYear, 'solMonth': self.solMonth}
        self.data = []
        self.info = []
        self.info2 = []

    def get_request_query(self, url, operation, params, serviceKey):
        import urllib.parse as urlparse
        params = urlparse.urlencode(params)
        request_query = url + '/' + operation + '?' + params + '&' + 'serviceKey' + '=' + serviceKey
        return request_query

    def getRequest(self):
        self.solYear = time.YEAR
        self.solMonth = str(time.MONTH).zfill(2)
        self.PARAMS = {'solYear': self.solYear, 'solMonth': self.solMonth}

        request_query = self.get_request_query(self.URL, self.OPERATION, self.PARAMS, self.SERVICEKEY)
        self.response = requests.get(url=request_query)
        print(self.response.json())
        try:
            for first in rqst.response.json().get('response').get('body').get('items').get('item'):
                self.data.append(int(datetime.datetime.strptime(str(first.get('locdate')), '%Y%m%d').day))
                self.info.append(first.get('dateName'))
                self.info2.append(first.get('isHoliday'))
        except AttributeError:
            print('Do nothing')

class Time:
    def __init__(self):
        self.datetime = QDateTime.currentDateTime()
        self.YEAR = int(self.datetime.toString('yyyy'))
        self.MONTH = int(self.datetime.toString('MM'))
        self.DAY = int(self.datetime.toString('dd'))
        self.weekdays = []
        self.days = 0

    def getDate(self):
        self.datetime = QDateTime.currentDateTime()
        self.YEAR = int(self.datetime.toString('yyyy'))
        self.MONTH = int(self.datetime.toString('MM'))
        self.DAY = int(self.datetime.toString('dd'))
        print(self.datetime)
        print(self.MONTH)

    def getWeekend(self):
        rqst.getRequest()
        for i in range(calendar.monthrange(self.YEAR, self.MONTH)[1]):
            wd = datetime.date(self.YEAR, self.MONTH, i+1).weekday()
            if wd == 5 or wd == 6 or (i+1) in rqst.data:
                self.weekdays.append('휴일')
            else:
                self.weekdays.append('미출근')


class TelBot:
    def __init__(self):
        self.telgm_token = 'token'
        self.bot = telegram.Bot(token=self.telgm_token)
        self.chatId = '-1001413179476'
        self.updates = self.bot.getUpdates()
        self.msgs = len(self.updates)

    def botCommand(self):
        self.updates = self.bot.getUpdates()
        #print(self.updates)
        if self.updates[-1] is None and self.updates[-1].message.chat.id == self.chatId:
            del self.updates[-1]
        elif len(self.updates) > self.msgs:
            msgs = len(self.updates)
            text = self.updates[-1].message.text
            if text == '안녕':
                self.sendMsg(self.updates[-1].message.chat.id, '안녕하세요')
            elif text in gdrive.members:
                msg = text + '님의 ' + gdrive.actSheet.cell(1, 1).value + ' 출근시간 입니다.\n'
                cell = gdrive.actSheet.find(text)
                cells_work = gdrive.actSheet.row_values(cell.row)
                cells_day = gdrive.actSheet.row_values(1)

                for i in range(1, time.DAY + 1):
                    msg = msg + cells_day[i] + ' ' + cells_work[i] + '\n'
                self.sendMsg(self.updates[-1].message.chat.id, msg)
            self.msgs = len(self.updates)

    def sendMsg(self, chatId, msg):
        self.bot.sendMessage(chat_id=chatId, text=msg)


class GDrive:
    def __init__(self):
        self.spreadsheet_url = 'url'
        self.scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive',
        ]
        self. dict = {
            "type": "service_account",
            "project_id": "worktime-265901",
            "private_key_id": "key",
            "private_key": "-----BEGIN PRIVATE KEY-----\nkey\n",
            "client_email": "mail",
            "client_id": "id",
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "auth_provider_x509_cert_url": "url",
            "client_x509_cert_url": "url"
        }
        self.credentials = ServiceAccountCredentials.from_json_keyfile_dict(self.dict, self.scope)
        self.gc = gspread.authorize(self.credentials)
        self.wb = self.gc.open_by_url(self.spreadsheet_url)
        self.actSheet = self.wb.worksheet(str(time.MONTH) + '월')
        self.memberSheet = self.wb.worksheet('members')
        self.setSheet = self.wb.worksheet('setting')
        self.WORKTIME = int(self.setSheet.cell(1, 2).value)
        self.members = self.memberSheet.col_values(1)

    def getActSheet(self):
        self.actSheet = self.wb.worksheet(str(time.MONTH) + '월')

    def newSheet(self):
        print('make new Sheet')
        time.getWeekend()
        print(len(time.weekdays))
        print(time.weekdays)
        print(rqst.info)
        print(rqst.info2)
        numOfMembers = len(self.members)
        lastday = calendar.monthrange(time.YEAR, time.MONTH)[1]

        self.actSheet = self.wb.add_worksheet(title=(str(time.MONTH) + '월'), rows=numOfMembers + 5, cols=lastday + 1)
        self.actSheet.update_cell(1, 1, self.actSheet.title)

        cell_list = self.actSheet.range(2, 1, numOfMembers + 1, 1)
        for i in range(len(cell_list)):
            cell_list[i].value = self.members[i]
        self.actSheet.update_cells(cell_list)

        cell_list = self.actSheet.range(1, 2, 1, lastday + 1)
        for i in range(len(cell_list)):
            cell_list[i].value = str(i + 1) + '일'
        self.actSheet.update_cells(cell_list)

        for i in range(1, numOfMembers + 1):
            cell_list = self.actSheet.range(i + 1, 2, i + 1, lastday + 1)
            for j in range(len(cell_list)):
                cell_list[j].value = time.weekdays[j]
            self.actSheet.update_cells(cell_list)

    def getMembers(self):
        self.members = self.memberSheet.col_values(1)

telgram = TelBot()
time = Time()
gdrive = GDrive()
rqst = Request()

if __name__ == "__main__":
    sched = BackgroundScheduler()
    sched.add_job(time.getDate, IntervalTrigger(seconds=1), id='test')
    sched.add_job(gdrive.getMembers, IntervalTrigger(hours=1), id='getMembers')
    sched.add_job(gdrive.newSheet, CronTrigger(day=1, hour=0, minute=0, second=3), id='newSheet')
    sched.start()

    while(True):
        telgram.botCommand()