# -*- coding: utf-8 -*-

import re
import pdfplumber
import math
from datetime import datetime, timedelta

# 记录上课、下课时分的字典
Begin_Hour = {'1': 8, '2': 9, '3': 10, '4': 11, '5': 14,
              '6': 15, '7': 16, '8': 17, '9': 19, '10': 19, '11': 20}
Begin_Minute = {'1': 30, '2': 20, '3': 20, '4': 10, '5': 30,
                '6': 20, '7': 10, '8': 0, '9': 0, '10': 50, '11': 40}
End_Hour = {'1': 9, '2': 10, '3': 11, '4': 11, '5': 15,
            '6': 16, '7': 16, '8': 17, '9': 19, '10': 20, '11': 21}
End_Minute = {'1': 2, '2': 0, '3': 0, '4': 50, '5': 10,
              '6': 0, '7': 50, '8': 40, '9': 40, '10': 30, '11': 20}


Year_Begin = '202x'  # 开学年份

Month_Begin = 'xx'  # 开学月份

Day_Begin = 'xx'  # 开学日

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
            days=self.weekday+7*(self.begin_week-1), hours=Begin_Hour[self.begin_section], minutes=Begin_Minute[self.begin_section])

        self.end_time = datetime(eval(Year_Begin), eval(Month_Begin.lstrip('0')), eval(Day_Begin.lstrip('0'))) + timedelta(
            days=self.weekday+7*(self.begin_week-1), hours=End_Hour[self.end_section], minutes=End_Minute[self.end_section])


class Caldenlar:
    lesson_header = ("课程名", "开始节次", "结束节次", "周数", "地点", "教师")

    def __init__(self, path):
        self.path = path
        self.total_lessons = []
        self.all_str_msg = []

    def __find_lesson_name_by_font_size(self, lessons_in_file, all_chars):
        for lessons_in_row in lessons_in_file:
            lessons_in_row = [
                i for i in lessons_in_row if i is not None and i != '']
            for lessons_in_table in lessons_in_row:
                first_lesson_regex = re.search(
                    r"(.*?[&#*])\n+\(\d+-\d+节\)", lessons_in_table, re.S)
                if first_lesson_regex is not None:
                    first_lesson_name = first_lesson_regex.group(
                        1).replace('\n', '')
                    str_msg = ''
                    title_font_size = 0
                    for i in range(len(all_chars)-1):
                        if title_font_size != 0 and all_chars[i]['size'] != title_font_size:
                            continue
                        if str_msg == '':
                            str_msg = all_chars[i]['text']
                        if all_chars[i]['size'] == all_chars[i+1]['size']:
                            str_msg += all_chars[i+1]['text']
                        else:
                            if title_font_size != 0:
                                self.all_str_msg.extend(
                                    re.split(r'[&#*]', str_msg))
                            elif str_msg == first_lesson_name:
                                title_font_size = all_chars[i]['size']
                                self.all_str_msg.extend(
                                    re.split(r'[&#*]', str_msg))
                            str_msg = ''
                    break
            else:
                continue
            break
        self.all_str_msg = list(set(self.all_str_msg))
        self.all_str_msg.remove('')

    def find_lesson(self):
        doc = pdfplumber.open(self.path)
        lessons_in_file = [
            lesson for page in doc.pages for lesson in page.extract_table()]

        # 去除无关紧要的行信息
        rows_deleted = []
        for i in range(len(lessons_in_file)):
            if len(lessons_in_file[i]) <= 1 or lessons_in_file[i][1] is None:
                rows_deleted.append(i)
                continue
            if not lessons_in_file[i][1].isdigit():
                rows_deleted.append(i)
            # 解决pdf分页而导致的数据断层问题
            if lessons_in_file[i][1] == '':
                # 找到本行中需要合并的列编号
                col_to_merge = []
                for j in range(len(lessons_in_file[i])):
                    if lessons_in_file[i][j] != '':
                        col_to_merge.append(j)
                # 往回搜索前几行，找到被割裂的课的上半部分，并与下半部分合并
                for j in range(i-1, 0, -1):
                    if lessons_in_file[j][col_to_merge[0]]:
                        for k in col_to_merge:
                            # pdf分页处会没有换行符信息，进而无法正确识别课程名字
                            lessons_in_file[j][k] += '\n'+lessons_in_file[i][k]
                        break
        # 删除无课程信息的行
        for row in rows_deleted[::-1]:
            lessons_in_file.pop(row)

        # 根据字体大小找出所有课程名
        self.__find_lesson_name_by_font_size(lessons_in_file, doc.chars)

        # 解包将二维列表转向，这样ics文件内课程顺序会按一周的时间顺序排（而不是按节次）
        lessons_in_file = list(zip(*lessons_in_file))
        # 剔除无用信息
        lessons_in_file.pop(0)
        lessons_in_file.pop(0)

        for weekday, lessons_in_row in enumerate(lessons_in_file):
            # 找出课程相关信息
            lessons_in_row = [
                i for i in lessons_in_row if i is not None and i != '']
            for lessons_in_table in lessons_in_row:
                lessons_in_table = re.findall(
                    r"(.*?)[&#*]\n+\((\d+)-(\d+)节\)(.*?周[^/]*)/(.*?)/(.*?)/(?:[^&#*]*?/)*.*?\n", lessons_in_table+'\n', re.S)
                lessons_in_table = [dict(zip(self.lesson_header, lesson))
                                    for lesson in lessons_in_table]
                # 一个表格内出现多门课程的情况
                for i in range(len(lessons_in_table)):
                    lesson_in_table = lessons_in_table[i]
                    # 更正课程名字，与上一课程的相关信息分开
                    if i != 0:
                        if lesson_in_table['课程名'].find('\n') != -1:
                            for j in self.all_str_msg:
                                if lesson_in_table['课程名'].find(j) != -1:
                                    lesson_in_table['课程名'] = j
                                    break

                    # 更正周数
                    # 出现 , 表示该课程在本学期出现的周数并不连续
                    all_weeks = lesson_in_table['周数'].split(
                        ',') if ',' in lesson_in_table['周数'] else [lesson_in_table['周数']]

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
                        # 去除换行符
                        for header in lesson_in_table:
                            if isinstance(lesson_in_table[header], str):
                                lesson_in_table[header] = lesson_in_table[header].replace(
                                    '\n', '')
                        new_lesson = Lesson(lesson_in_table)
                        self.total_lessons.append(new_lesson)

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
    print('课表ics文件已生成')
