try:
    from tty_menu import tty_menu
except ImportError:
    from util import input_number

    def tty_menu(menu, title= None):
        print(title)
        i = 1
        for m in menu:
            print('%d. %s' % (i, m))
            i += 1
        pos = input_number('请输入序号：', min=1, max=len(menu))
        return int(pos) - 1

if __name__ == '__main__':
    menus = [
        ('🔄 刷新', lambda: True),
        ('🔄 重新加载文件', lambda: True),
        ('📄 生成一个文件', lambda: True),
        ('📧 发送一封邮件', lambda: True),
        ('👻 生成全部文件和邮件（不发送）', lambda: True),
        ('📨 全部发送', lambda: True),
        ('📨 发送某个序号之后', lambda: True),
        ('📨 只发送离职的和没写邮箱的', lambda: True),
        ('🚪 退出程序', exit),
    ]
    pos = tty_menu([t[0] for t in menus], "请选择?")
    print(menus[pos])