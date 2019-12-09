import wikipedia
import datetime
import re

def get_today_anniversary():
    dt_now = datetime.datetime.now()
    today = dt_now.strftime('%-m月%-d日')
    wikipedia.set_lang('ja')
    words = wikipedia.search(today)
    page = wikipedia.page(words[0]).content

    pattern = r'(==\s記念日・年中行事\s==.*?)\n+==\s.*\s=='
    m = re.search(pattern, page, flags=re.DOTALL)
    anniv = m.groups()[0].replace('。', '。\n')
    title = '= ' + today + ' =\n'
    return title + anniv

def main():
    anniv = get_today_anniversary()
    print(anniv)

if __name__ == '__main__':
    main()
