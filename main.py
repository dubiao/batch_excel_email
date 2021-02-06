
import time
import progressbar

from cofing_DEFAULT import DEFAULT_CONF
from emailer import Emailer
from util import select_file, input_number, input_email, input_month
from attach_gen import EmailGenerator
from configuration import Configuration
from excel_reader import SalaryFileReader
from menu_wrap import tty_menu
import excel_reader
import subprocess


def get_user_info_by_index(excel: SalaryFileReader):
    seq = input_number('请输入员工序号: ', exit='exit')
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


def input_and_execute(excel: SalaryFileReader, generator: EmailGenerator, emailer: Emailer) :
    seq = input_number('请输入序号: ', exit='exit')
    if seq is None:
        return None
    execute_all(excel, generator, emailer, seq)



def execute_all(excel: SalaryFileReader, generator: EmailGenerator, emailer: Emailer, start_seq: int = 1):
    bye_people_need_email = True
    bye_people_has_other_email = False
    no_email_need_email = True
    only_bye_no_email_people = False
    if emailer:
        if excel.has_noemail_people:
            how_for_byes = [
                '不发送工资邮件',
                '手动输入邮箱',
            ]
            if tty_menu(how_for_byes, "发现表格中有部分成员没有邮箱。对于他们将?") == 0:
                no_email_need_email = False
        if excel.has_bye_people:
            how_for_byes = [
                '按表格中的邮箱发送',
                '不发送工资邮件',
                '手动输入离职人员邮箱',
            ]
            ipos = tty_menu(how_for_byes, "发现表格中有离职成员。对于他们将?")
            if ipos == 0:
                bye_people_need_email = True
                bye_people_has_other_email = False
            elif ipos == 1:
                bye_people_need_email = False
                bye_people_has_other_email = False
            elif ipos == 2:
                bye_people_need_email = True
                bye_people_has_other_email = True
    if start_seq < 0:
        only_bye_no_email_people = True
    with progressbar.ProgressBar(max_value=len(excel.user_value_map.items()) + 1, redirect_stdout=True, ) as p:
        p.update(0)
        for i, user_info in excel.user_value_map.items():
            p.update(i)
            # user_info = excel.user_value_map[i]
            path = generator.make_file(user_info, make=True)
            if user_info['seq'] < start_seq:
                continue
            if only_bye_no_email_people and (user_info['email'] and not user_info['out_day'] ):
                continue

            if emailer and user_info['out_day']:
                p.finish(dirty=True)
                print('%s已于%s离职。' % (user_info['name'], user_info['out_day']))
                if bye_people_need_email:
                    if bye_people_has_other_email:
                        user_info['email'] = input_email("请输入邮箱 或 直接回车 %s:\n" % (('使用' + user_info['email']) if user_info['email'] else '不发送'), default=user_info['email'], empty=True)
                    else:
                        if not user_info['email'] and no_email_need_email:
                            user_info['email'] = input_email("请输入邮箱:\n")
                        pass
                else:
                    user_info['email'] = None
            elif emailer and user_info['email'] is None:
                p.finish(dirty=True)
                print('%s未填写邮箱。' % (user_info['name']))
                if no_email_need_email:
                    user_info['email'] = input_email("请输入邮箱:\n", empty=True)
            if user_info['email'] is None:
                print('%s 未发送邮件。文件保存在：%s\n' % (user_info['name'], path))
                continue
            # p.start()
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
    emailer = Emailer()
    config = Configuration(CONF_PATH, DEFAULT_CONF)
    ym = input_month()
    file_path = select_file('请选择文件')
    # ym = {'year': 2020, 'month': 12}
    # file_path = 'assets/%d%02d.xlsx' % (ym['year'], ym['month'])
    if 'file_sheet_name_fmt' not in config:
        config['file_sheet_name_fmt'] = '{month}'

    generator = EmailGenerator(config['generate_file'], **ym)
    sheet_name = config['file_sheet_name_fmt'].format(month_code=generator.for_month_code)
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
        ('🔄 刷新', lambda: print(reader.table(config['table_view_head']))),
        ('🔄 重新加载文件', 'read_xlsx'),
        ('📄 生成一个文件', lambda: open_named_files_by_indexes(reader, generator)),
        ('📧 发送一封邮件', lambda: send_email_by_indexes(config, reader, generator, emailer)),
        ('👻 生成全部文件和邮件（不发送）', lambda: execute_all(reader, generator, None)),
        ('📨 全部发送', lambda: execute_all(reader, generator, emailer)),
        ('📨 发送某个序号之后', lambda: input_and_execute(reader, generator, emailer)),
        ('📨 只发送离职的和没写邮箱的', lambda: execute_all(reader, generator, emailer, -1)),
        ('🚪 退出程序', exit),
    ]

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

