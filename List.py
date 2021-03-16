# -*- coding: utf-8 -*-

import re
import pdfplumber
import math
from datetime import datetime, timedelta

# 记录上课、下课时分的字典
TimeTable = {
    "石牌": {
        'Begin_Hour':   {'1': 8,  '2': 9,  '3': 10, '4': 11, '5': 14, '6': 15, '7': 16, '8': 17, '9': 19, '10': 19, '11': 20},
        'Begin_Minute': {'1': 30, '2': 20, '3': 20, '4': 10, '5': 30, '6': 20, '7': 10, '8': 0,  '9': 0,  '10': 50, '11': 40},
        'End_Hour':     {'1': 9,  '2': 10, '3': 11, '4': 11, '5': 15, '6': 16, '7': 16, '8': 17, '9': 19, '10': 20, '11': 21},
        'End_Minute':   {'1': 10, '2': 0,  '3': 0,  '4': 50, '5': 10, '6': 0,  '7': 50, '8': 40, '9': 40, '10': 30, '11': 20}
    },
    "大学城": {
        'Begin_Hour':   {'1': 8,  '2': 9,  '3': 10, '4': 11, '5': 14, '6': 14, '7': 15, '8': 16, '9': 19, '10': 19, '11': 20},
        'Begin_Minute': {'1': 30, '2': 20, '3': 20, '4': 10, '5': 0,  '6': 50, '7': 40, '8': 30, '9': 0,  '10': 50, '11': 40},
        'End_Hour':     {'1': 9,  '2': 10, '3': 11, '4': 11, '5': 14, '6': 15, '7': 16, '8': 17, '9': 19, '10': 20, '11': 21},
        'End_Minute':   {'1': 10, '2': 0,  '3': 0,  '4': 50, '5': 40, '6': 30, '7': 20, '8': 10, '9': 40, '10': 30, '11': 20}
    },
    "南海": {
        'Begin_Hour':   {'1': 8,  '2': 9,  '3': 10, '4': 11, '5': 14, '6': 14, '7': 15, '8': 16, '9': 19, '10': 19, '11': 20},
        'Begin_Minute': {'1': 30, '2': 20, '3': 20, '4': 10, '5': 0,  '6': 50, '7': 40, '8': 30, '9': 0,  '10': 50, '11': 40},
        'End_Hour':     {'1': 9,  '2': 10, '3': 11, '4': 11, '5': 14, '6': 15, '7': 16, '8': 17, '9': 19, '10': 20, '11': 21},
        'End_Minute':   {'1': 10, '2': 0,  '3': 0,  '4': 50, '5': 40, '6': 30, '7': 20, '8': 10, '9': 40, '10': 30, '11': 20}
    }
}


Year_Begin = '202x'  # 开学年份

Month_Begin = 'xx'  # 开学月份

Day_Begin = 'xx'  # 开学日

Location = "xx"  # 石牌 / 大学城 / 南海

Pdf_Path = 'xxx(20xx-20xx-x)课表.pdf'  # 相对路径或绝对路径

Trigger_Time = 20  # 提前多少分钟提醒（默认为20分钟，此数值不建议调成太大）


class Lesson:

    def __init__(self, info_lesson):
        self.name = info_lesson['课程名']
        self.place = info_lesson['地点']
        self.teacher = info_lesson['教师']
        self.begin_week = int(info_lesson['开始周数'])
        self.lasting_time = info_lesson['持续次数']
        self.begin_section = info_lesson['开始节次']
        self.end_section = info_lesson['结束节次']
        self.weekday = info_lesson['星期几']
        self.odd_dual = True if info_lesson['单双周'] is not None else False

        self.begin_time = datetime(eval(Year_Begin), eval(Month_Begin.lstrip('0')), eval(Day_Begin.lstrip('0'))) + timedelta(
            days=self.weekday+7*(self.begin_week-1), hours=TimeTable[Location]['Begin_Hour'][self.begin_section], minutes=TimeTable[Location]['Begin_Minute'][self.begin_section])

        self.end_time = datetime(eval(Year_Begin), eval(Month_Begin.lstrip('0')), eval(Day_Begin.lstrip('0'))) + timedelta(
            days=self.weekday+7*(self.begin_week-1), hours=TimeTable[Location]['End_Hour'][self.end_section], minutes=TimeTable[Location]['End_Minute'][self.end_section])


class Caldenlar:
    lesson_header = ('周数', '地点', '教师', '课程名')

    def __init__(self, path):
        self.path = path
        self.total_lessons = []

    def find_lesson(self):
        doc = pdfplumber.open(self.path)
        lessons_in_file = []
        for page in doc.pages:
            lessons_in_file.extend(page.extract_table())
        weekday = -1

        # 剔除无效信息
        lessons_in_file = [
            row for row in lessons_in_file if row[-1] is not None]

        for row_index in range(len(lessons_in_file)):
            if("星期" in str(lessons_in_file[row_index][0])):
                weekday += 1

            lesson_in_table = re.search(
                r"周数:\s*(.*?)\s*地点:\s*(.*?)\s*教师:\s*(.*?)\s.*?\n(.*?)[*&#]\n", lessons_in_file[row_index][-1], re.S)

            lesson_in_table = dict(zip(self.lesson_header,
                                       lesson_in_table.groups()))
            if lessons_in_file[row_index][1] is not None:
                lesson_in_table['开始节次'], lesson_in_table['结束节次'] = lessons_in_file[row_index][1].split(
                    '-')
            else:
                # 回溯前面的行寻找课程时间信息
                for back_row_index in range(row_index-1, -1, -1):
                    if lessons_in_file[back_row_index][1] is not None:
                        lesson_in_table['开始节次'], lesson_in_table['结束节次'] = lessons_in_file[back_row_index][1].split(
                            '-')
                        break

            # 出现 , 表示周数在一个学期并不连续
            if ',' in lesson_in_table['周数']:
                all_weeks = lesson_in_table['周数'].split(',')
            else:
                all_weeks = [lesson_in_table['周数']]

            for j in range(len(all_weeks)):
                week_split = re.search(
                    r'(\d+)(?:-(\d+))?周\n?(?:\(\n?([单双])\n?\))?', all_weeks[j], re.S)
                lesson_in_table['开始周数'], lesson_in_table['结束周数'], lesson_in_table['单双周'] = week_split.groups(
                )
                if lesson_in_table['结束周数'] is None:
                    lesson_in_table['持续次数'] = 1
                    # 出现单双周情况，则重复次数减少将近一半
                elif lesson_in_table['单双周'] is not None:
                    lesson_in_table['持续次数'] = math.ceil((
                        eval(lesson_in_table['结束周数']) - eval(lesson_in_table['开始周数']) + 1)/2)
                else:
                    lesson_in_table['持续次数'] = eval(
                        lesson_in_table['结束周数']) - eval(lesson_in_table['开始周数']) + 1
                lesson_in_table['星期几'] = weekday

                # 为对象变量赋值
                newLesson = Lesson(lesson_in_table)
                self.total_lessons.append(newLesson)

    def produce_lesson(self):
        text = 'BEGIN:VCALENDAR\n'
        for lesson in self.total_lessons:
            text += 'BEGIN:VEVENT\n'
            text += 'SUMMARY:{0}（{1}）\n'.format(lesson.name, lesson.teacher)
            text += 'DTSTART:{0}\n'.format(
                lesson.begin_time.strftime("%Y%m%dT%H%M00"))
            text += 'DTEND:{0}\n'.format(
                lesson.end_time.strftime("%Y%m%dT%H%M00"))
            text += 'LOCATION:{0}\n'.format(lesson.place)
            text += 'RRULE:FREQ=WEEKLY;COUNT={0};'.format(
                lesson.lasting_time)
            # 如果涉及到单双周问题，则加入INTERVAL选项，每两周重复一次
            if lesson.odd_dual:
                text += 'INTERVAL=2;'
            text += '\nBEGIN:VALARM\nTRIGGER:-P0DT0H{}M0S\nEND:VALARM\n'.format(
                Trigger_Time)
            text += 'END:VEVENT'+'\n'
        text += 'END:VCALENDAR'
        fileName = self.path[:-3]+'ics'
        with open(fileName, 'w', encoding='utf-8') as f:
            f.write(text)


if __name__ == '__main__':
    lesson_in_term = Caldenlar(Pdf_Path)
    lesson_in_term.find_lesson()
    lesson_in_term.produce_lesson()
    print("课表ics文件已生成")
