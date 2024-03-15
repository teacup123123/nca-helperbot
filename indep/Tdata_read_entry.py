"""閱讀 資料夾中所有pdf 產出xlsx表格"""
import os
import re
from collections import OrderedDict
from datetime import datetime
from pypdf import PdfReader


def run(data_folder):
    full = []
    for fn in os.listdir(data_folder):
        if not fn.endswith('.pdf'): continue
        reader = PdfReader(fn)
        page = reader.pages[0]
        txt = page.extract_text()
        lines = txt.splitlines()

        contenti = 0
        key = ''
        entry = OrderedDict()

        for li, line in enumerate(lines):
            if line == '：':
                value = ''.join(lines[contenti:li - 1])
                contenti = li + 1
                if key:

                    filt_regex_date = r'\d+/\d+/\d+'
                    maps_reformat_dates = [
                        lambda x: x.split('/'),
                        lambda x: f'{int(x[0]) + 1911}/{x[1]}/{x[2]}',
                        lambda x: datetime.strptime(x, '%Y/%m/%d').date(),
                    ]

                    filt = {k: filt_regex_date for k in '列印時間 法定期滿日 預退日期 入營日期 出生日期'.split()}
                    maps = {k: maps_reformat_dates for k in '列印時間 法定期滿日 預退日期 入營日期 出生日期'.split()}
                    # maps['印表人'] = lambda v: v[:3]
                    if key in filt:
                        value = re.match(filt[key], value).group(0)
                    if key in maps:
                        for map in maps[key]:
                            value = map(value)

                    entry[key] = value
                key = lines[li - 1]

        # finally
        value = ''.join(lines[contenti:li + 1])

        entry[key] = value
        full.append(entry)
        # print(entry)

    # print(page.extract_text())

    keys = [k for k in full[0].keys() if all(k in entry.keys() for entry in full)]
    assert '梯次' in keys
    assert all(entry['梯次'] == full[0]['梯次'] for entry in full)
    Tnumber = full[0]['梯次']

    print(keys)

    if SAVE := False:
        from openpyxl import Workbook

        wb = Workbook()
        # grab the active worksheet
        ws = wb.active
        ws.title = Tnumber
        # Data can be assigned directly to cells
        ws.append(keys)
        for entry in full:
            print(entry)
            ws.append([entry[k] for k in keys])

        wb.save(f"Tdata_compiled/database{Tnumber}.xlsx")

        # time is acceptable!
        # ws['A2'] = datetime.datetime.now()



if __name__ == '__main__':
    run('data253')
