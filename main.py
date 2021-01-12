
import time
import progressbar
from tty_menu import tty_menu

from cofing_DEFAULT import DEFAULT_CONF
from emailer import Emailer
from util import select_file, input_number, input_email, input_month
from attach_gen import EmailGenerator
from configuration import Configuration
from excel_reader import SalaryFileReader
import excel_reader
import subprocess


def get_user_info_by_index(excel: SalaryFileReader):
    seq = input_number('请输入序号: ', exit='exit')
    if seq is None:
        return None
    else:
        return excel.get_info(seq)


def open_named_files_by_indexes(excel: SalaryFileReader, generator: EmailGenerator):
    user_info = get_user_info_by_index(excel)
    path = generator.make_file(user_info, )
    print(path)
    # os.startfile(path)
    subprocess.call(["open", path])
    pass


def send_email_by_indexes(config: Configuration, excel: SalaryFileReader, generator: EmailGenerator, emailer: Emailer):
    email = config['tester_email']
    user_info = get_user_info_by_index(excel)
    if user_info:
        email = input_email('请输入邮箱: ' if email is None else '请输入邮箱 或直接使用 %s: ' % email, False, email)
        if email:
            config['tester_email'] = email
        else:
            email = config['tester_email']
        if not email:
            return False
        path = generator.make_file(user_info)
        print(path)
    else:
        print('Error index')
        return False
    print('---->', email)
    email_content = generator.make_email(user_info)
    return emailer.send(email, email_content['subject'], email_content['content'], path)
    pass


def execute_all(excel: SalaryFileReader, generator: EmailGenerator, emailer: Emailer):
    bye_people_need_email = True
    bye_people_has_other_email = False
    if emailer and excel.has_bye_people:
        how_for_byes = [
            '按表格中的邮箱发送',
            '不发送工资邮件',
            '手动输入离职人员邮箱',
        ]
        pos = tty_menu(how_for_byes, "发现表格中有离职成员。对于他们将?")
        if pos == 0:
            bye_people_need_email = True
            bye_people_has_other_email = False
        elif pos == 1:
            bye_people_need_email = False
            bye_people_has_other_email = False
        elif pos == 2:
            bye_people_need_email = True
            bye_people_has_other_email = True

    with progressbar.ProgressBar(max_value=len(excel.user_value_map.items()) + 1, redirect_stdout=True, ) as p:
        for i, user_info in excel.user_value_map.items():
            # user_info = excel.user_value_map[i]
            path = generator.make_file(user_info, make=False)
            if emailer and user_info['out_day']:
                print('%s已于%s离职。' % (user_info['name'], user_info['out_day']))
                if bye_people_need_email:
                    if bye_people_has_other_email:
                        user_info['email'] = input_email("\n请输入邮件 或 直接使用 %s:\n" % user_info['email'], default=user_info['email'])
                    else:
                        pass
                else:
                    print('不发送邮件。文件保存在：%s' % path)
                    user_info['email'] = None
                    continue

            print('---->', user_info['email'])
            email_content = generator.make_email(user_info)
            if emailer:
                emailer.send(user_info['email'], email_content['subject'], email_content['content'], path)
            else:
                print(email_content['subject'])
                print(email_content['content'])
            p.update(i + 1)
            time.sleep(0.05)
    pass


CONF_PATH = 'conf.json'

if __name__ == '__main__':
    config = Configuration(CONF_PATH, DEFAULT_CONF)
    ym = input_month()
    file_path = select_file('请选择文件')
    # ym = {'year': 2020, 'month': 12}
    # file_path = 'assets/%d%02d.xlsx' % (ym['year'], ym['month'])
    if 'file_sheet_name_fmt' not in config:
        config['file_sheet_name_fmt'] = '{month}'

    generator = EmailGenerator(config['generate_file'], **ym)
    sheet_name = config['file_sheet_name_fmt'].format(month=generator.for_month_code)
    reader = SalaryFileReader(file_path, sheet_name, config['template_title_map'])
    try:
        reader.load()
    except excel_reader.SheetNotFound as e:
        # 加载容错
        if len(e.sheet_names):
            pos = tty_menu([t for t in e.sheet_names], "请问数据在哪个Sheet中：")
            reader.set_sheet_name(e.sheet_names[pos])
            reader.load()
    except ValueError:
        raise ValueError
        pass

    print(reader.table(config['table_view_head']))

    menus = [
        ('刷新', lambda: print(reader.table(config['table_view_head']))),
        ('生成一个文件', lambda: open_named_files_by_indexes(reader, generator)),
        ('发送一封邮件', lambda: send_email_by_indexes(config, reader, generator, emailer)),
        ('生成全部文件和邮件（不发送）', lambda: execute_all(reader, generator, None)),
        ('重新加载文件', 'read_xlsx'),
        ('全部发送', lambda: execute_all(reader, generator, emailer)),
        ('退出程序', exit),
    ]
    emailer = Emailer()
    while True:
        pos = tty_menu([t[0] for t in menus], "请选择?")
        if pos is None:
            exit(0)
        if pos < len(menus):
            func = menus[pos][1]
            if func == 'read_xlsx':
                reader.load()
                print(reader.table(config['table_view_head']))
            else:
                if len(menus[pos]) > 2 and menus[pos][2]:
                    func(config, reader, generator, emailer)
                else:
                    func()

