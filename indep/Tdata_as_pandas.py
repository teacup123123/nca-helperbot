"""xlsx表格 轉換成pandas資料結構 存入二進位檔.pickle"""
import os
import re

import pandas as pd

full = None
for fn in os.listdir('Tdata_compiled'):
    if re.match('database\d+.xlsx', fn):
        read = pd.read_excel(os.path.join('Tdata_compiled', fn))
        read.to_pickle(os.path.join('Tdata_compiled', fn.replace('xlsx', 'pickle')))
    full: pd.DataFrame = read if full is None else pd.concat([full, read])
full.to_pickle(os.path.join('Tdata_compiled', 'allT.pickle'))
