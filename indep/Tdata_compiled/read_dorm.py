import os
import re

import openpyxl as ox
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

here, _ = os.path.split(__file__)


def read_dorm(src=os.path.join(here, '役政司役男宿舍分配一覽1130229.xlsx')):
    res = {}
    workbook: Workbook = ox.load_workbook(src)
    for ws in workbook.worksheets:
        for row in ws.rows:
            for cell in row:
                if (
                        cell.value is not None and
                        isinstance(cell.value, str) and
                        (m_ := re.match(r'(\S{2}\S?)\n?\d\d?/\d\d?', cell.value))
                ):
                    print(f'{m_.group(1)} belongs to {ws.title[:2]}')
                    res[m_.group(1)] = ws.title[:2]
    return res


if __name__ == '__main__':
    read_dorm()
