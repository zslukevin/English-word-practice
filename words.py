#!/usr/bin/env python
# coding=utf-8
# **********************************************************************
# Time    : 2018/9/20 19:00
# Author  : kevin lu
# Function: 
# Python  : 3.6 
# **********************************************************************
import sqlite3
import os, time, re
from random import shuffle
from typing import Dict, Any
from reportlab.pdfgen import canvas
from reportlab.lib import pagesizes, units, colors
from reportlab.pdfbase import ttfonts, pdfmetrics


def getUnits():
    db = sqlite3.connect(os.getcwd() + '/high_school_words_data/words.db')
    cursor = db.cursor()
    sql = 'select 单元 from words group by 单元'
    result = cursor.execute(sql)
    unit_list = []
    for row in result.fetchall():
        unit_list.append(row[0])
    cursor.close()
    db.close()
    return unit_list


class WordsPrint:
    __records: Dict[Any, Any]

    def __init__(self, page_config=None):
        self.__db = sqlite3.connect(os.getcwd() + '/high_school_words_data/words.db')
        self.__cursor = self.__db.cursor()
        self.__id_list = []
        self.__records = {}
        self.__page_conf = page_config
        self.__start_unit = '1.1'
        self.__end_unit = '1.1'
        default_page_conf = {
            'pagesize': pagesizes.A4,
            'columns': 2,
            'col_padding': 5 * units.mm,
            'left': 10 * units.mm,
            'top': 15 * units.mm,
            'right': 10 * units.mm,
            'bottom': 15 * units.mm,
            'line_height': 10 * units.mm
        }
        self.__page_conf = default_page_conf.copy()
        if type(page_config) is dict:
            for item in default_page_conf:
                if item in page_config:
                    if item == 'pagesize' and hasattr(pagesizes, page_config['pagesize'].upper()):
                        self.__page_conf['pagesize'] = eval('pagesizes.' +  page_config['pagesize'].upper())
                    elif item in ['col_padding', 'left', 'top', 'right', 'bottom', 'line_height']:
                        self.__page_conf[item] = page_config[item] * units.mm
                    else:
                        self.__page_conf[item] = page_config[item]

    def setRange(self, start_unit, end_unit, content_conf, is_Rand=True):
        content_list = content_conf['list']
        self.__start_unit = start_unit
        self.__end_unit = end_unit
        sql = 'select id, 单元, 类型, 单词, 词性, 解析 from words where 单元 >= :start and 单元 <= :end and 类型 in (' + content_list + ')'
        if content_conf['onlyError']:
            _isEng = '英文' if content_conf['isEnglish'] else '中文'
            sql = sql + f' and {_isEng}错次 > 0'
        result = self.__cursor.execute(sql, {'start': start_unit, 'end': end_unit})
        records = result.fetchall()
        self.__db.commit()
        self.__id_list = []
        self.__records = {}
        for row in records:
            self.__id_list.append(row[0])
            self.__records[row[0]] = {
                    '单元': row[1].strip(),
                    '类型': row[2].strip(),
                    '单词': row[3].strip(),
                    '词性': '' if row[4] is None else row[4].strip(),
                    '解析': row[5].strip()
            }
            
        if is_Rand:
            shuffle(self.__id_list)

    def loadTest(self, t_id):
        sql = f'select 卷名, 生成时间, 类型, 题列表, 内容 from testings where ID=:id'
        result = self.__cursor.execute(sql, {'id': t_id})
        t_info =result.fetchone()
        if len(t_info)<1:
            return False
        print(t_info)
        t_name = t_info[0]
        t_time = t_info[1]
        t_type = t_info[2]
        t_list = t_info[3]
        t_conent = t_info[4]
        t_isEng = False
        t_onlyError = False
        sql = f'select id, 单元, 类型, 单词, 词性, 解析 from words where ID in ({t_list})'
        result = self.__cursor.execute(sql)
        records = result.fetchall()
        self.__db.commit()
        self.__id_list = []
        self.__records = {}
        for row in records:
            self.__records[row[0]] = {
                '单元': row[1].strip(),
                '类型': row[2].strip(),
                '单词': row[3].strip(),
                '词性': '' if row[4] is None else row[4].strip(),
                '解析': row[5].strip()
            }
        pattern = re.compile(r'([(].*?[)])')
        t_match = re.search(pattern, t_name)
        t_unit = ''
        if t_match.group(1):
            t_unit = t_match.group(1)[1:-1]
            t_range = t_unit.split('-')
            self.__start_unit = t_range[0]
            self.__end_unit = t_range[1]
        t_match =re.search(r'(仅错题)', t_conent)
        if t_match:
            t_onlyError = True
        if t_type[0:3] == '看英文':
            t_isEng=True
        s_list = []
        for x in t_list.split(','):
            s_list.append(int(x))
        self.__id_list = s_list
        return (t_unit, t_isEng, t_onlyError)

            
    @property
    def getWords(self):
        return self.__id_list, self.__records

    def printPDF(self, file_name='words.pdf', isEng=True, answer=False, onlyError=False, d_page=False, msg_obj=None):
        page_conf = self.__page_conf
        pdfmetrics.registerFont(ttfonts.TTFont('SimSun', os.getcwd() + '/high_school_words_data/simsun.ttc'))
        page = canvas.Canvas(file_name, page_conf['pagesize'])
        w, h = page_conf['pagesize']
        top_padding = page_conf['top']
        bottom_padding = page_conf['bottom']
        left_padding = page_conf['left']
        right_padding = page_conf['right']
        line_height = page_conf['line_height']
        columns = page_conf['columns']
        col_padding = page_conf['col_padding']
        fix_width = w - left_padding - right_padding
        fix_height = h - top_padding - bottom_padding - line_height
        font_size = 12
        _ch_error = '错题' if onlyError else ''
        page_title = '高中英语第%s-%s单元单词(词组)%s练习[%s]' % (self.__start_unit, self.__end_unit, _ch_error, time.strftime('%Y-%m-%d'))

        page.translate(left_padding, h - top_padding)
        page.setTitle(page_title)
        page.setFont('SimSun', 8)
        page.drawCentredString(fix_width / 2, 0, page_title)
        page.setFont('SimSun', font_size)
        page.setStrokeColor(colors.black)

        col_width = fix_width / columns - col_padding
        lines = fix_height // line_height + 1
        i = 0
        c = 0
        list_count = len(self.__id_list)
        for id in self.__id_list:
            y = -(i % lines + 1) * line_height
            x = (c % columns) * (col_width + col_padding)
            this_word = self.__records[id]
            text_line = this_word['单词'] if isEng else this_word['解析']
            if this_word['类型'] == '单词':
                text_line = text_line + '(' + this_word['词性'] + ')'
            text_line = text_line.strip()
            text_length = len(text_line.encode('gbk'))
            l_start = x + text_length * font_size / 2
            l_end = x + col_width
            page.drawString(x, y, text_line)
            page.line(l_start, y - 3, l_end, y - 3)
            if (i+1) % lines == 0:
                c = c + 1
                if c % columns == 0:
                    page.setFont('SimSun', 8)
                    page.drawCentredString(fix_width//2, -fix_height-line_height, '- %d -' % page.getPageNumber())
                    page.showPage()
                    page.translate(left_padding, h - top_padding)
                    page.setFont('SimSun', 8)
                    page.drawCentredString(fix_width // 2, 0, page_title)
                    page.setFont('SimSun', font_size)
                    page.setStrokeColor(colors.black)
            i = i + 1
            if not msg_obj is None:
                msg_obj.showMessage('正在生成题面...{:.2f}%'.format(i/list_count*100))
        page.setFont('SimSun', 8)
        page.drawCentredString(fix_width//2, -fix_height-line_height, '- %d -' % page.getPageNumber())
        if d_page and page.getPageNumber() % 2 > 0:
            page.showPage()
        the_page = page.getPageNumber()
        page.showPage()
        if answer:
            page_title = '高中英语第%s-%s单元单词(词组)%s练习答案[%s]' % (self.__start_unit, self.__end_unit, _ch_error, time.strftime('%Y-%m-%d'))
            page.translate(left_padding, h - top_padding)
            page.setTitle(page_title)
            page.setFont('SimSun', 8)
            page.drawCentredString(fix_width / 2, 0, page_title)
            page.setFont('SimSun', font_size)
            page.setStrokeColor(colors.black)
            i = 0
            c = 0
            for id in self.__id_list:
                y = -(i % lines + 1) * line_height
                x = (c % columns) * (col_width+col_padding)
                this_word = self.__records[id]
                text_line = this_word['单词'] if isEng else this_word['解析']
                text_answer = this_word['解析'] if isEng else this_word['单词']
                if this_word['类型'] == '单词':
                    text_line = text_line + '(' + this_word['词性'] + ')'
                text_line = text_line.strip()
                text_length = len(text_line.encode('gbk'))
                l_start = x + text_length * font_size / 2
                l_end = x + col_width
                page.drawString(x, y, text_line+' '+text_answer)
                page.line(l_start, y - 3, l_end, y - 3)
                if (i + 1) % lines == 0:
                    c = c + 1
                    if c % columns == 0:
                        page.setFont('SimSun', 8)
                        page.drawCentredString(fix_width // 2, -fix_height - line_height,
                                               '- {:d} -'.format(page.getPageNumber() - the_page))
                        page.showPage()
                        page.translate(left_padding, h - top_padding)
                        page.setFont('SimSun', 8)
                        page.drawCentredString(fix_width // 2, 0, page_title)
                        page.setFont('SimSun', font_size)
                        page.setStrokeColor(colors.black)
                i = i + 1
                if not msg_obj is None:
                    msg_obj.showMessage('正在生成答案...{:.2f}%'.format(i / list_count * 100))
            page.setFont('SimSun', 8)
            page.drawCentredString(fix_width // 2, -fix_height - line_height,
                                   f'- {page.getPageNumber() - the_page:d} -')
            if d_page and page.getPageNumber() % 2 > 0:
                page.showPage()
            page.showPage()
        page.save()




if __name__ == '__main__':
    start_u = '1.2'
    end_u = '1.3'
    words = WordsPrint()
    words.setRange(start_u, end_u)
    id, rows = words.getWords
    for i in id:
        print(rows[i])
    words.printPDF()
    pageconf = {
        'pagesize': 'a3',
        'columns': 4,
        'left': 10
    }
    newpdf = WordsPrint(pageconf)
    newpdf.setRange('1.1', '1.5')
    newpdf.printPDF('newPDF.pdf')


