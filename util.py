import time
import re
import tkinter as tk
from tkinter import filedialog


def select_file(message: str):
    filepath = filedialog.askopenfilename(
        title=message,
        filetypes=[('Excel', '*.xlsx'), ('All Files', '*')],
        initialdir='.'
    )
    return filepath


def input_month():
    current = time.localtime()
    print('当前日期: ', time.strftime("%Y-%m-%d", current))
    guess_year = current.tm_year if current.tm_mon > 1 else current.tm_year - 1
    guess_month = current.tm_mon - 1 if current.tm_mon > 1 else 12
    year_month = input_number('请输入年月 或按回车确认 %d%d: ' % (guess_year, guess_month), True, 6, 202012, 222212)
    if not year_month:
        year_month = {
            'year': guess_year,
            'month': guess_month
        }
    else:
        year = int(year_month // 100)
        month = year_month % 100
        if month < 1 or month > 12:
            print('请去火星')
            exit(0)
        year_month = {
            'year': year,
            'month': month
        }
    return year_month


def input_number(message='请输入数字: ', empty=False, length=None, min=None, max=None, exit=None):
    while True:
        try:
            just_input = input(message)
            if just_input:
                if just_input == exit:
                    return None
                if length and len(just_input) != length:
                    continue
                value = int(just_input)
                if min is not None and min > value:
                    continue
                if max is not None and max < value:
                    continue
            elif empty:
                return None
            else:
                continue
        except ValueError:
            continue
        else:
            return value


def input_email(message='请输入邮箱: ', empty=False, default=None):
    empty_times = 0
    while True:
        try:
            just_input = input(message)
            if just_input:
                regex = re.compile(r'''(
[a-zA-Z0-9._%+-]+ # username
@ # @ symbol
([a-zA-Z0-9.-]+) # domain name
(\.[a-zA-Z]{2,4}) # dot-something
)''', re.VERBOSE)
                mo = regex.search(just_input)
                if mo:
                    return mo.group()
                else:
                    continue
                pass
            elif empty or default:
                if empty_times:
                    empty_times += 1
                    if empty_times > 2:
                        return None
                return default
            else:
                empty_times += 1
                if empty_times > 2:
                    return None
                continue
        except ValueError:
            continue
