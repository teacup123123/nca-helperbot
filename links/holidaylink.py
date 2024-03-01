from urllib.parse import quote,unquote



def genlink(formid,id='252T8888', name='王小喵', ):
    return (f'https://docs.google.com/forms/d/e/1FAIpQLSed6ECASUd0Gze1hHgykRqlZMOt3to7zQrMoStTb6SXF1645Q'
            f'/viewform?usp=pp_url'
            f'&entry.1813326080={formid}'
            f'&entry.286310740={id}'
            f'&entry.330405440={quote(name)}'
            f'&entry.927673877=glory'
            f'&entry.1599983468=reason'
            f'&entry.615287842=2024-03-01'
            f'&entry.325271870=0800'
            f'&entry.1459125740=2024-03-01'
            f'&entry.1614929291=1200'
            )
if __name__ == '__main__':
    # print(quote('王大明'))
    # print(unquote('%E5%BC%B5%E8%BF%AA%E5%87%B1'))
    print(genlink(12345,'252T8519','你看我可以預先寫好表單'))