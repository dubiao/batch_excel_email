import dload
import os
url = "https://github.com/dubiao/batch_excel_email/archive/master.zip"
if __name__ == '__main__':
    dload.save_unzip(url, '.', True)

    if not os.path.exists('assets'):
        os.mkdir('assets')

    if not os.path.exists('run.bat'):
        run_win = '''python batch_excel_email-master/main.py'''
        with open('run.bat', 'w') as run:
            run.writelines(run_win)

    if not os.path.exists('run.bat'):
        with open('update.bat', 'w') as run:
            run_win = '''python download_unzip.py'''
            run.writelines(run_win)
