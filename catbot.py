import collections
import dataclasses
import json
import os.path
# import pickle
import re
import signal
import sys
from datetime import datetime
from urllib.parse import quote, unquote

import requests
from linebot.v3.webhooks import MessageEvent

from credentials import line_chan_acctoken
from state_machine import bot, transition_from


class file_backed_dict(dict):
    def __init__(self, filename=None):
        super().__init__()
        self.filename = ''
        if filename:
            self.filename = filename
            with open(filename, 'r') as f:
                read = json.load(f)
                # read = pickle.load(f)
            for k, v in read.items():
                self[k] = catbot(**v)

    def save(self, filename=None):
        with open(filename, 'w') as f:
            copy = {}
            for k, v in self.items():
                copy[k] = dataclasses.asdict(v)
            json.dump(copy, f, indent=4)
            # pickle.dump(copy, f)
            # pickle.dump(self, f)

    def __setitem__(self, key, value):
        super().__setitem__(key, value)
        if self.filename:
            self.save(self.filename)


@dataclasses.dataclass
class catbot(bot):
    id: str = ''
    name: str = ''

    @transition_from('init')
    def prompt_for_name(self, query: str, **kwargs):
        match = re.match(r'(\D{2}\D?)T(\d{4})', query)
        if match and match.span()[0] == 0 and match.span()[1] == len(query):
            self.name, self.id = match.groups()
            return 'named', f'{self.name}大大好~'
        else:
            return 'init', "抱歉(cony cry)，我不知道你是誰，麻煩請符合格式自我介紹...如王大明T8123"

    @transition_from('named')
    def await_for_actions(self, query: str, **kwargs):
        if query == '其他，我有話要跟孝丞說，其實，一直以來...':
            return ('waiting for msg',
                    f"請說請說(moon halo)，我會幫你把訊息印出來，屬名{self.name}然後放小信封，用(heart)貼紙封好，放在他鞋盒裡。請耐心等候他的回信")
        elif query == '我要更新自介，手殘編號名字寫錯了':
            return ('init',
                    f"OK~ 麻煩請符合格式自我介紹...如王大明T8123")
        elif query == '我要請假，我要掃描紙本假單，後續給AI自己處理':
            return ('await holiday img',
                    f"請[紙本假單]照片上傳給我喔~(heart)")
        elif query == '我要申請[出差請示單]，我要掃描紙本公假單，後續給AI自己處理':
            return ('await outing img',
                    f"請[出差請示單]照片上傳給我喔~(heart)")
        elif query == '我要請假，我怕AI笨笨，我自己去填google表單':
            return ('named',
                    f"OK~ 表格在這裡，流水號記得寫紙本正面左上角喔: https://docs.google.com/forms/d/e/"
                    f"1FAIpQLSed6ECASUd0Gze1hHgykRqlZMOt3to7zQrMoStTb6SXF1645Q/viewform?usp=pp_url"
                    f"&entry.1813326080={datetime.now().strftime('%m%d-%H%M%S')}"
                    f"&entry.286310740={self.id}"
                    f"&entry.330405440={quote(self.name)}"
                    )
        else:
            return 'named', f'{self.name}大大，拍謝~(cony cry)我聽不懂你說的"{query}"建議點擊預設動作'

    @transition_from('await holiday img')
    def await_holiday_img(self, query: str, **kwargs):
        got = requests.get(f'https://api-data.line.me/v2/bot/message/{kwargs["msgid"]}/content',
                           headers={"Authorization": f"Bearer {line_chan_acctoken}"}, stream=True)
        filename = datetime.now().strftime('%m%d-%H%M%S')
        with open(os.path.join('holiday_forms', filename + '.jpg'), 'wb') as f:
            for chunk in got:
                f.write(chunk)

        return 'named', '收到檔案了，我的血輪眼[Google Vision]還沒練好，所以先存孝丞那邊'
        # do some Google Vision magic here TODO
        return 'named',
        (f"根據我過目(moon sunglasses)(moon sunglasses)，表格幫你預填了，流水號記得寫紙本正面左上角喔: "
         f"https://docs.google.com/forms/d/e/"
         f"1FAIpQLSed6ECASUd0Gze1hHgykRqlZMOt3to7zQrMoStTb6SXF1645Q/viewform?usp=pp_url"
         f"&entry.1813326080={filename}"
         f"&entry.286310740={self.id}"
         f"&entry.330405440={quote(self.name)}")

    @transition_from('await outing img')
    def await_outing_img(self, query: str, **kwargs):
        got = requests.get(f'https://api-data.line.me/v2/bot/message/{kwargs["msgid"]}/content',
                           headers={"Authorization": f"Bearer {line_chan_acctoken}"}, stream=True)
        filename = datetime.now().strftime('%m%d-%H%M%S')
        with open(os.path.join('holiday_forms', filename + '.jpg'), 'wb') as f:
            for chunk in got:
                f.write(chunk)
        return 'named', '收到檔案了，我的血輪眼[Google Vision]還沒練好，所以先存孝丞那邊'
        # do some Google Vision magic here TODO
        return 'named',
        (f"根據我過目(moon sunglasses)(moon sunglasses)，表格幫你預填了，流水號記得寫紙本正面左上角喔: "
         f"https://docs.google.com/forms/d/e/"
         f"1FAIpQLSed6ECASUd0Gze1hHgykRqlZMOt3to7zQrMoStTb6SXF1645Q/viewform?usp=pp_url"
         f"&entry.1813326080={filename}"
         f"&entry.286310740={self.id}"
         f"&entry.330405440={quote(self.name)}")


    @transition_from('waiting for msg')
    def await_msg(self, query: str, **kwargs):
        now = datetime.now()  # current date and time
        filename = now.strftime("%Y.%m.%d-%H.%M.%S({}).txt").format(self.name)
        with open(os.path.join('shoebox', filename), 'w', encoding='utf-8') as f:
            f.write(query)
        return ('named',
                "放到鞋盒裡啦~(heart)(heart)")


# if INIT_BLANK := True:
#     initial = file_backed_dict()
#     initial.save('secrets/users.txt')

userbase = file_backed_dict('secrets/users.txt')  # usrid->catbot


def signal_handler(sig, frame):
    print('You pressed Ctrl+C! will save database')
    userbase.save('secrets/users.pickle')
    sys.exit(0)


signal.signal(signal.SIGINT, signal_handler)

print('catbot successfully loaded')
