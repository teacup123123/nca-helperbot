from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from datetime import date as Date
import numpy as np
from openpyxl import Workbook
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.worksheet import Worksheet
import openpyxl as ox
import re
from datetime import date as Date, datetime as Datetime
from enum import Enum
import update_monthly_ledger as uml
from interactions import interaction, start_interaction
from update_monthly_ledger import workhours, grab_till
# import prettytable as pt
from tabulate import tabulate
import sys

import pandas as pd

target = 'holiday_ledger/組改期間役男榮譽假清冊3_改.xlsx'
wb_global: Worksheet = None

full_data = pd.read_pickle('Tdata_compiled/allT.pickle')


def rows_no_header(ws):
    header = None
    for row in ws.rows:
        if header is None:
            header = row
            continue
        yield tuple(c for c in row)


def id2name(uid):
    uid = int(uid)
    # res: pd.DataFrame = full_data[full_data.uid.isin([uid-8000, uid, 8000 + uid])]
    res: pd.DataFrame = full_data[full_data.uid == uid % 8000]
    assert len(res) == 1
    return res.name.values[0]


def obtain_name(uid_or_name):
    if isinstance(uid_or_name, int):
        uid = int(uid_or_name)
        name = id2name(uid)
    elif re.match('\d{3}\d?', uid_or_name):
        uid = int(uid_or_name)
        name = id2name(uid)
    else:
        name = uid_or_name
    return name


def sort_toppriority(ws):
    rvals = []
    for row in rows_no_header(ws):
        rvals.append(tuple(x.value for x in row))

    def key(x):
        exp = tuple(int(xx) for xx in re.findall('\d+', x[7])) if x[7] is not None else (111, 1, 1)
        comp = x[0] is not None, x[5] - x[6] == 0, exp, x[5] - x[6]  # 預價先, 有空先, 快過期者先, 剩餘較少者優先
        return comp

    rvals.sort(key=key)

    for row, rowv in zip(rows_no_header(ws), rvals):
        for cell, val in zip(row, rowv):
            cell.value = val


@interaction
def use_holiday(uid_or_name, timestring: str):
    """

    :param uid_or_name: 123 or 8123 or 王大明
    :param timestring: 1130325 0830-1130326 1330
    :return:
    """
    timestring = timestring.replace('-', '~')
    timestring = re.sub('(1\d{2})(\d{2})(\d{2})', '\g<1>/\g<2>/\g<3>', timestring)

    name = obtain_name(uid_or_name)

    if wb_global is None: load()
    ws: Worksheet = wb_global[name]

    fr, to = grab_till(timestring)
    yr, mn, dy, frtime = fr
    yr2, mn2, dy2, totime = to
    usage = sum(
        Datetime.combine(Date(yr + 1911, mn, dy), frtime) <=
        x <=
        Datetime.combine(Date(yr2 + 1911, mn2, dy2), totime) for x in uml.workhours
    )
    decompress = (f'{yr}/{mn}/{dy} {frtime.hour:02}{frtime.minute:02}'
                  f'~{yr2}/{mn2}/{dy2} {totime.hour:02}{totime.minute:02} 自動特')

    denom = usage
    while usage != 0:
        sort_toppriority(ws)
        if ws['F2'].value - ws['G2'].value == 0:
            ws.append((None,) * 3 + ('預假', None, usage, usage, ws['H2'].value, decompress + f'({usage}/{denom})'))
            print(f'{name} 的 "{decompress}" 需要預假 {usage}/{denom}')
            sort_toppriority(ws)
            break
        use = min(ws['F2'].value - ws['G2'].value, usage)
        usage = usage - use
        line = f'{decompress}({use}/{denom})'
        print(f'{name} {ws["D2"].value}(~{ws["H2"].value};{ws["J2"].value}) "" 添加 "{line}"')
        if ws['I2'].value is None: ws['I2'].value = ''
        ws['I2'].value = '\n'.join(ws['I2'].value.splitlines() + [line])
        ws['G2'].value += use
    refresh_format(name)



@interaction
def refresh_format(uid_or_name):
    name = obtain_name(uid_or_name)

    if wb_global is None: load()
    ws: Worksheet = wb_global[name]

    for r in rows_no_header(ws):
        for c in r:
            c.font = Font()
            c.border = Border()
            c.alignment = Alignment()


@interaction
def save_new(postfix):
    wb_global.save(target.replace('.xlx', postfix + '.xlsx'))


@interaction
def save():
    wb_global.save(target)


@interaction
def load():
    global wb_global
    wb_global = ox.load_workbook(target)


@interaction
def print_sheet(uid_or_name):
    name = obtain_name(uid_or_name)
    if wb_global is None: load()
    ws = wb_global[name]
    rows = []
    for r in rows_no_header(ws):
        rows.append([c.value for c in r])

    print(tabulate(rows, headers=[ws[f'{c}1'].value for c in 'ABCDEFGHIJ'], tablefmt='fancy_grid'))


if __name__ == '__main__':
    if sys.argv:
        with open(sys.argv[1], 'r') as f:
            lines = f.readlines()


        def mygen():
            for l in lines: yield l.strip()


        g = mygen()


        def myinput(prompt=''):
            try:
                n = next(g)
                print(prompt, end=f'{n}\n')
                return n
            except StopIteration:
                return input(prompt)


        start_interaction(myinput)
    else:
        start_interaction()
    # id2name(519)
    # id2name(8519)
    use_holiday(519, '1130315 0900-19 1700')
    # append_new_holiday(519, '1130315 0900-19 1700(8)')
    refresh_format(519)
    print()
