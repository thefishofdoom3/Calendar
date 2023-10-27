from __future__ import print_function
import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
from oauth2client import file

import gspread
from google.oauth2.service_account import Credentials

from datetime import datetime, timedelta


class RegistrationInfo:
    def __init__(self, data):
        self.email = data[1]
        self.name = data[2]
        self.initialSchedule = data[3]
        self.schedule = data[4]
        self.keyForVLookUp = data[5]
        self.sessionCount = int(data[6].split()[0])
        self.nomination = data[7] == 'Yes'
        self.shortDate = data[8]
        self.courseName = data[9]
        self.loca = data[10]

class CourseInfo:
    def __init__(self, data):
        self.competency = data[0]
        self.nameEN = data[1]
        self.nameVI = data[2]
        self.objectiveEN = data[3]
        self.objectiveVI = data[4]

class CourseTrainer:
    def __init__(self, data):
        self.courseName = data[4]
        self.PICOwner = data[6]
        self.zoomLink = data[21]

class GroupObject:
    def __init__(self, key, courseInfo, courseTrainer):
        self.key = key
        firstSplits = key.rsplit('(',1)
        self.time = firstSplits[1].rsplit(')', 1)[0] + ', ' + firstSplits[0].rsplit('] ')[1]
        self.isOnline = "ONLINE" in key
        self.registrationInfos = []
        self.courseInfo = courseInfo
        self.courseTrainer = courseTrainer

    def addRegistrationInfo(self, registrationInfo):
        self.registrationInfos.append(registrationInfo)

class Manager:
    def __init__(self):
        self.groupObjects = []
        self.courseInfos = []
        self.courseTrainers = []

    def getCourseInfoFromList(self, courseName):
        for obj in self.courseInfos:
            if (obj.nameEN.lower() == courseName.lower()):
                return obj

    def getCourseTrainerFromList(self, courseName):
        for obj in self.courseTrainers:
            if (obj.courseName.lower() == courseName.lower()):
                return obj
            
    def getGroupObjectFromList(self, key):
        for obj in self.groupObjects:
            if (obj.key.lower() == key.lower()):
                return obj
            
    def addCourseInfo(self, courseInfo):
        exist = self.getCourseInfoFromList(courseInfo.nameEN)
        if exist is None:
            self.courseInfos.append(courseInfo)

    def addCourseTrainer(self, courseTrainer):
        exist = self.getCourseTrainerFromList(courseTrainer.courseName)
        if exist is None:
            self.courseTrainers.append(courseTrainer)

    def addRow(self, registrationInfo):
        exist = self.getGroupObjectFromList(registrationInfo.schedule)

        if exist is None:
            courseInfo = self.getCourseInfoFromList(registrationInfo.courseName)
            courseTrainer = self.getCourseTrainerFromList(registrationInfo.courseName)
            if courseInfo is None or courseTrainer is None:
                return
            groupObject = GroupObject(registrationInfo.schedule, courseInfo, courseTrainer)
            groupObject.addRegistrationInfo(registrationInfo)
            self.groupObjects.append(groupObject)
        else:
            exist.addRegistrationInfo(registrationInfo)


# CREATE EVENT ON CALENDAR

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'credentials.json'  #credentials json get from your google account 
APPLICATION_NAME = 'Google Calendar API Python Quickstart'


def convert_to_bullet_points(text): #just text editing function :>
    i = 1
    sentences = text.split('\n')
    bullet_points = []
    
    for sentence in sentences:
        if not sentence.strip():
            continue
        if i == 1:
            bullet_points.append(f"   • {sentence[2:].strip()}")
        else:
            bullet_points.append(f"       • {sentence[2:].strip()}")

        i +=1
    
    bullet_point_text = '\n'.join(bullet_points) 
    return bullet_point_text


def convert_datetime(input_datetime): #input format '14:00 - 16:00, Ngày 1 tháng 8' to get format '2023-08-01T09:00:00-07:00'

    time_range_str = input_datetime.split(',')[0].strip()

    start_time_str, end_time_str = time_range_str.split(' - ')

    start_time = datetime.strptime(start_time_str, '%H:%M')
    end_time = datetime.strptime(end_time_str, '%H:%M')

    date_str = input_datetime.split(',')[1].strip()
    day_str, month_str = date_str.split(' tháng ')

    day = int(day_str.split()[-1])
    month = int(month_str)

    current_year = datetime.now().year

    start_datetime = datetime(current_year, month, day, start_time.hour, start_time.minute)

    end_datetime = start_datetime + (end_time - start_time)

    time_zone_offset = timedelta(hours=7)

    formatted_start = (start_datetime + time_zone_offset).strftime('%Y-%m-%dT%H:%M:%S%z') + '+07:00'
    formatted_end = (end_datetime + time_zone_offset).strftime('%Y-%m-%dT%H:%M:%S%z') + '+07:00'

    return formatted_start, formatted_end


def get_credentials():

    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def main():


    number = 1
    # Set up Google Sheets API credentials
    sheet_credentials = Credentials.from_service_account_file('ps.json') #json of your google service account
    scoped_credentials = sheet_credentials.with_scopes(
    ['https://www.googleapis.com/auth/spreadsheets'])

    # Authorize and create the client
    client = gspread.authorize(scoped_credentials)

    sheet = client.open_by_url(
    'sheet_link_1').worksheet('Sheet1')
    # Get the parameter values from specific cells
    data = sheet.get_all_values()
    # sheet menu khoá học
    sheet_menu = client.open_by_url(
    'sheet_link_2').worksheet('Sheet2')
    data_menu = sheet_menu.get_all_values()
    # sheet để mapping giảng viêng
    sheet_tutor = client.open_by_url(
    'sheet_link_3').worksheet('Sheet3')
    data_tutor = sheet_tutor.get_all_values()


    manager = Manager()

    for i in range(len(data_menu)):
        if i > 2:
            courseInfo = CourseInfo(data_menu[i])
            manager.addCourseInfo(courseInfo)

    for i in range(len(data_tutor)):
        if i > 2:
            courseTrainer = CourseTrainer(data_tutor[i])
            manager.addCourseTrainer(courseTrainer)

    for i in range(len(data)):
        if i > 0:
            registrationInfo = RegistrationInfo(data[i])
            manager.addRow(registrationInfo)

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    for i in range(len(manager.groupObjects)):
        
        groupObject = manager.groupObjects[i]


        if groupObject.isOnline:
            
            # TEMPLATE EVENT ONLINE
            html_template_onl = """

    Chào bạn,

<b> Thân mời bạn tham gia lớp <b>{vnName} ({engName}) chi tiết như sau. </b>
    
<b> 1. Thông tin buổi đào tạo

        • Thời gian: {time}
        • Zoom Link: {link}
        • Ngôn ngữ: Tiếng Việt
        • Giảng viên: {tutorName} </b>

<b> 2. Mục tiêu buổi đào tạo</b>

    {obj}

<b> 3. Chuẩn bị trước buổi đào tạo </b>

    content

<b> Nội quy lớp học:</b>
    content

<b> Lưu ý:</b>

    content
"""
            params = {
                'vnName': manager.groupObjects[i].courseInfo.nameVI,
                'engName': manager.groupObjects[i].courseInfo.nameEN,
                'time': manager.groupObjects[i].time,
                'link': manager.groupObjects[i].courseTrainer.zoomLink,
                'tutorName': manager.groupObjects[i].courseTrainer.PICOwner,
                'obj': convert_to_bullet_points(manager.groupObjects[i].courseInfo.objectiveEN),
                'time2': manager.groupObjects[i].time,
                'link_1': 'link'
            }
            formatted_html = html_template_onl.format(**params)
        else:
# TEMPLATE EVENT OFFLINE

            html_template_off = """

    Chào bạn,

    Thân mời bạn tham gia lớp <b>{vnName} ({engName}) chi tiết như sau. </b>
    
<b> 1. Thông tin buổi đào tạo

        • Thời gian: {time}
        • Ngôn ngữ: Tiếng Việt
        • Giảng viên: {tutorName} </b>

<b> 2. Mục tiêu buổi đào tạo</b>

    {obj}

<b> 3. Chuẩn bị trước buổi đào tạo </b>

        • Đọc trước một số nội dung tại đây (Pre-read/Pre-survey) trước {time2}
        content

<b> Lưu ý dành cho lớp Offline:</b>
        • Vui lòng đọc kỹ hướng dẫn tham gia lớp đào tạo trực tiếp <a href="{link_2}">tại đây</a>.

"""
            params_2 = {
                'vnName': manager.groupObjects[i].courseInfo.nameVI,
                'engName': manager.groupObjects[i].courseInfo.nameEN,
                'time': manager.groupObjects[i].time,
                'link': manager.groupObjects[i].courseTrainer.zoomLink,
                'tutorName': manager.groupObjects[i].courseTrainer.PICOwner,
                'obj': convert_to_bullet_points(manager.groupObjects[i].courseInfo.objectiveEN),
                'time2': manager.groupObjects[i].time,
                'link_1': 'link1',
                'link_2': 'link2'
            }


            # Substitute the placeholders with parameter values
            formatted_html = html_template_off.format(**params_2)
        # print (formatted_html)


        attendees = []
        for registrationInfo in groupObject.registrationInfos:
            attendees.append({'email': registrationInfo.email})

        # print(attendees)
        # return
        location = data[i][10]
        eng_name = manager.groupObjects[i].courseInfo.nameEN
        start_time, end_time = convert_datetime(manager.groupObjects[i].time)
        time = manager.groupObjects[i].time

        event = {
            'summary': f'G-Calendar Title: [{location} Class] Shopee Academy | {eng_name}| {time}',
            'location': '',
            'description': formatted_html,
            'start': {
                'dateTime': f'{start_time}',
                'timeZone': 'Asia/Ho_Chi_Minh',
            },
            'end': {
                'dateTime': f'{end_time}',
                'timeZone': 'Asia/Ho_Chi_Minh',
            },
            'recurrence': [
                'RRULE:FREQ=DAILY;COUNT=2'
            ],
            'attendees': attendees,
            'reminders': {
                'useDefault': False,  # Disable default reminders
                'overrides': [
                    {'method': 'popup', 'minutes': 1440},
                    {'method': 'email', 'minutes': 1440}],  
            },
        }

        try:
            event = service.events().insert(calendarId='primary', body=event, sendUpdates='all').execute()
            print (f'đây là event thứ {number}')
            print(manager.groupObjects[i].courseInfo.nameEN)
            print('Event created: %s' % (event.get('htmlLink')))
            number +=1 
        except Exception as e:
            print('Error creating event:', e)


if __name__ == '__main__':

    main()
