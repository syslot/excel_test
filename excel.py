#!/Users/ningyu/Source/virtualenv/excel/bin/python
# coding=utf-8
from __future__ import division

import pandas as pd
import os
import time
import sys
import platform
import requests
import click
from openpyxl import load_workbook

sheet_list = [
    "3-1",
    "3-2",
    "3-3",
    "3-4",
    "3-5",
    "3-6",
    "3-7",
    "3-8",
    "3-9",
    "4-1",
    "4-2",
    "4-3",
    "4-4",
    "5-1",
    "5-2",
    "5-3",
    "5-4",
    "5-5",
    "5-6",
    "5-7",
    "5-8",
    "5-9",
    "5-10",
    "5-11",
    "5-12",
]

audio_dir = '/Users/ningyu/Desktop/audio/'
p_version = platform.system()
word_new = ["word"]

src_excel = '/Users/ningyu/Desktop/listening/record copy.xlsx'
dst_excel = '/Users/ningyu/Desktop/listening/foo.xlsx'


# mode schema 0x            a                   b
#             ^--- hex      ^--- filter index   ^--- filter relation
# 'a' = 0 ,filter all indexes
# 'a' = 0 ,filter last two indexes
# 'b' = 1 ,filter indexes filtered by 'or' relation
# 'b' = 0 ,filter indexes filtered by 'and' relation
def filter_sheet(sheetname, mode):
    ps = pd.read_excel(src_excel, sheetname=sheetname)
    tmp_list = []
    if (mode >> 1) & 0x1 == 0:
        for i in ps.columns[1:]:
            if ps[i].count() == 2:
                continue
            tmp_list.append(u'(ps[\'%s\']==u\'✕\')' % i)
    elif (mode >> 1) & 0x1 == 1:
        for i in ps.columns[-2:]:
            if ps[i].count() == 2:
                continue
            tmp_list.append(u'(ps[\'%s\']==u\'✕\')' % i)
    if (mode & 0x01) == 1:
        newps = eval('ps[' + '|'.join(tmp_list) + ']')
    else:
        newps = eval('ps[' + '&'.join(tmp_list) + ']')

    word_list = newps['word'].tolist()
    word_new.extend(word_list)


def read_sheet(delay=5):
    ps = pd.read_excel(dst_excel, sheetname="Sheet1")
    word_list = ps["word"].tolist()
    for word in word_list:
        if p_version == 'Darwin':
            os.system('say ' + word)
        elif p_version == 'Windows':
            print(audio_dir + word.replace(' ', '') + '.mp3')
            os.system(audio_dir + word.replace(' ', '') + '.mp3')
        else:
            print("Not support")
        if ' ' in word:
            delay *= 2
        time.sleep(delay)


def cal(df):
    # df=pd.DataFrame
    df["word"].values[-4] = df["word"].count() - 1
    for i in df.columns[1:]:
        df[i].values[-4] = df[i].count()
        df[i].values[-1] = 1 - df[i].values[-4] / df["word"].values[-4]


def merge_sheet(sheetname):
    ps = pd.read_excel(src_excel, sheetname=sheetname)
    ps_new = pd.read_excel(dst_excel, sheetname="Sheet1")
    df = pd.merge(ps, ps_new, on="word", how="left")
    cal(df)
    book = load_workbook(src_excel)
    writer = pd.ExcelWriter(src_excel, engine='openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    df.to_excel(writer, sheetname, index=False)
    writer.save()


def send_request(word):
    # My API
    # GET http://dict.youdao.com/dictvoice

    try:
        response = requests.get(
            url="http://dict.youdao.com/dictvoice",
            params={
                "type": "0",
                "audio": word,
            },
        )
        return response.content
    except requests.exceptions.RequestException:
        print('HTTP Request failed')


# check for audio
def get_audio(sheetname, mode):
    ps = pd.read_excel(src_excel, sheetname=sheetname)
    word_list = ps['word'].tolist()
    for word in word_list[:-4]:
        path = audio_dir + word.replace(' ', '') + '.mp3'
        if os.path.isfile(path):
            continue

        ctx = send_request(word)
        wf = open(path, "wb")
        wf.write(ctx)
        wf.close()


@click.command()
@click.option(
    '--work',
    default='filter',
    help='excel.py working mode, filter&get_audio&merge&speed,filter params: sheet,mode,loop, merge params: sheet,loop, read params: speed')
@click.option(
    '--sheet',
    default=0,
    help='filter mode begin sheet,default = 3-1')
@click.option('--mode', default=0, help='filter rules, default = all error')
@click.option('--loop', default=5, help='filter sheet sum')
@click.option('--speed', default=1.5, help='Speed of reading')
def main(work, sheet, mode, loop, speed):
    if work == 'get_audio':
        for sheet in sheet_list:
            get_audio(sheet, '0')

    all_len = len(sheet_list)
    sheet_index = sheet

    # Dictation Mode
    if work == 'filter':
        for i in range(loop):
            filter_sheet(sheet_list[(sheet_index + i) % all_len], mode)

        # save dst excel
        pf = pd.DataFrame(word_new)
        pf.to_excel(dst_excel, sheet_name='Sheet1', index=False, header=False)

        # read_sheet(speed)

    # Merge sheet
    if work == 'merge':
        for i in range(loop):
            merge_sheet(sheet_list[(sheet_index + i) % all_len])

    # Read Excel
    if work == 'read':
        read_sheet(speed)


if __name__ == '__main__':
    main()
