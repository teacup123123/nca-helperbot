"""自動填寫 月份假冊 補修榮譽價欄位"""
import collections
import re
from collections import defaultdict
import openpyxl as ox
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from datetime import date as Date, timedelta, datetime as Datetime, time as Time

month_repre = Date(2024, 3, 15)  # 目標月份，取月中+-28日之請假事項


def gen_workdays():
    start_date = Date(2024, 1, 1)  # TODO 所有工作日
    end_date = Date(2025, 12, 12)  # TODO 所有工作日
    load_holidays = [  # TODO 減去法定假日
        'holiday_ledger/113年中華民國政府行政機關辦公日曆表_Google行事曆專用.csv'
    ]

    import csv
    set_holidays = set()
    work_days = []
    for csvfile in load_holidays:
        with open(csvfile, 'r',
                  encoding='utf8') as holidays:
            reader = csv.reader(holidays)
            for i, row in enumerate(reader):
                if i == 0: continue  # header
                _date = row[1]
                if _date == '':
                    continue
                # set_holidays.add(_date)
                set_holidays.add(Datetime.strptime(_date, '%Y/%m/%d').date())
                # set_holidays.add(date.fromisoformat(_date.replace('/','-')))

    def daterange(start_date, end_date):
        for n in range(int((end_date - start_date).days)):
            yield start_date + timedelta(n)

    for single_date in daterange(start_date, end_date):
        if single_date not in set_holidays:
            # print(single_date.strftime("%Y-%m-%d"))
            work_days.append(single_date)
    return work_days


def gen_hours(_workdays=None):
    if _workdays is None:
        _workdays = workdays
    workhours = []
    for d in _workdays:
        for _i in range(4):
            workhours.append(Datetime.combine(d, Time(8 + _i, 45)))
            workhours.append(Datetime.combine(d, Time(13 + _i, 31)))

    # New year later hours removed like this TODO 特殊非全天放假
    for _ in [
        Datetime(2024, 2, 7, 15, 31),
        Datetime(2024, 2, 7, 16, 31)
    ]:
        if _ in workhours: workhours.remove(_)

    byworkday = collections.defaultdict(list)
    for _ in workhours: byworkday[_.date()].append(_)
    return workhours, byworkday


workdays = gen_workdays()
workhours, byworkday = gen_hours()

pattern_date3 = r'(\d{3})/(\d+)/(\d+)'
pattern_detectfrac = r'\(([\d?]+)/([\d?]+)\)'


def grab_till(token: str):
    timefr, timeto = 0, 2359
    token = re.match('[\d/~\-\s]+\d', token).group(0)
    _ = token.replace('-', '~').split('~')
    if len(_) == 2:
        fr, till = _
        fr = fr.replace('/', ' ').split()
        till = till.replace('/', ' ').split()
        if re.match('\d{4}', fr[-1]):
            timefr = int(fr[-1])
            fr.pop()
        if re.match('\d{4}', till[-1]):
            timeto = int(till[-1])
            till.pop()
    else:
        fr, *till = _
        fr = fr.replace('/', ' ').split()
        if re.match('\d{4}', fr[-1]):
            timefr = int(fr[-1])
            fr.pop()
    fr = fr + [timefr]
    till = till + [timeto]
    while len(till) != len(fr):
        till = [fr[-1 - len(till)]] + till
    fr, till = list(map(int, fr)), list(map(int, till))
    fr[-1] = Time(fr[-1] // 100, fr[-1] % 100)
    till[-1] = Time(till[-1] // 100, till[-1] % 100)
    return fr, till


if __name__ == '__main__':
    wbsrc: Workbook = ox.load_workbook('holiday_ledger/組改期間役男榮譽假清冊3_改.xlsx')
    skipsheet = True
    seen = defaultdict(set)
    for worksheet in wbsrc.worksheets:  # worksheet 即分頁
        if worksheet.title == '北辦': continue  # 不弄北辦
        if worksheet.title != '張瀚文' and skipsheet:
            continue
        else:
            skipsheet = False
        # 會從張瀚文一直弄到北辦前

        name = worksheet.title

        for cell in worksheet['I'][1:]:  # I是使用紀錄，1是跳掉開頭的"使用紀錄"
            if cell.value is None: continue
            content: str = cell.value
            for line in content.splitlines():  # each line is one (partial usage of holiday)
                # but may be distributed over many sources
                # we must remove duplicates, we do this via the dict(name->set) 'seen'
                if line.startswith('*預'): continue  # already used, we skip
                token, n, yr2 = re.search('([\w\W]+)' + pattern_detectfrac, line).groups()
                yr, mn, dy = re.search(pattern_date3, token).groups()
                date = tuple(int(x) for x in (yr, mn, dy))  # to number
                seen[name].add((date, token, int(yr2)))
    print(seen)

    wbdst: Workbook = ox.load_workbook(f'holiday_ledger/{month_repre.month}月請假清冊統整 孝丞製作_自動.xlsx')
    worksheet: Worksheet = wbdst.active  # only one sheet
    for _i, row in enumerate(worksheet.rows):
        if _i == 0: continue  # skip header
        name = row[2].value
        if not name in seen: continue

        enter_into = row[4]

        day_counter = defaultdict(int)
        for (yr, mn, dy), token, hrs in seen[name]:
            if abs(month_repre - Date(1911 + yr, mn, dy)) > timedelta(days=28): continue

            # a, b, c = re.search(pattern_date3, token).groups()
            _fr, _to = grab_till(token)
            yr, mn, dy, starttime = _fr
            yr2, mn2, dy2, endtime = _to
            yr = int(yr) + 1911
            yr2 = int(yr2) + 1911
            yr, mn, dy, yr2, mn2, dy2 = map(int, [yr, mn, dy, yr2, mn2, dy2])
            # starttime = re.search('\d{4}', token)
            # if starttime: starttime = starttime.group(0)

            if any(x in token for x in '早四 晚四 當日'.split()) or not any(x in token for x in '-~'):
                usage = hrs
                if mn == month_repre.month:
                    day_counter[yr, mn, dy] += usage
                # NO VERIFICATION!!! because 公益大使團 就地較複雜
            elif '全日八' in token:
                usage = len(byworkday[Date(yr, mn, dy)])
                assert hrs == usage
                if mn == month_repre.month:
                    day_counter[yr, mn, dy] += usage
            elif '多日' in token:
                print(f'SPECIAL CARE for {token}')
                totusage = 0
                for dd in filter(lambda ddd: Date(yr, mn, dy) <= ddd <= Date(yr2, mn2, dy2), workdays):
                    usage = sum(
                        1 for x in byworkday[dd] if
                        (
                                Datetime.combine(Date(yr, mn, dy), starttime) <= x <=
                                Datetime.combine(Date(yr2, mn2, dy2), endtime)
                        )
                    )
                    totusage += usage
                    if dd.month == month_repre.month:
                        day_counter[dd.year, dd.month, dd.day] += usage
                        # lines.append(f'{dd.year}/{dd.month:02}/{int(dd.day):02}:{dz}日{usage}時')
                        # print(lines[-1])
                    print()
                assert hrs == totusage
            else:
                print(f'SPECIAL CARE for {token}')
                totusage = 0
                for dd in filter(lambda ddd: Date(yr, mn, dy) <= ddd <= Date(yr2, mn2, dy2), workdays):
                    usage = sum(
                        1 for x in byworkday[dd] if
                        (
                                Datetime.combine(Date(yr, mn, dy), starttime) <= x <=
                                Datetime.combine(Date(yr2, mn2, dy2), endtime)
                        )
                    )
                    totusage += usage
                    if dd.month == month_repre.month:
                        day_counter[dd.year, dd.month, dd.day] += usage
                        # lines.append(f'{dd.year}/{dd.month:02}/{int(dd.day):02}:{dz}日{usage}時')
                        # print(lines[-1])
                    print()
                assert hrs == totusage

        lines = []
        for (yr, mn, dy), usage in day_counter.items():
            dz, usage = usage // 8, usage % 8
            lines.append(f'{yr}/{mn:02}/{dy:02}:{dz}日{usage}時')
        enter_into.value = '\n'.join(sorted(lines))

    wbdst.save(f'holiday_ledger/{month_repre.month}月請假清冊統整 孝丞製作_自動.xlsx')
