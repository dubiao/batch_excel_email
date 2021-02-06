import zipfile
import requests
import os
url = "https://github.com/dubiao/batch_excel_email/archive/master.zip"
if __name__ == '__main__':
    if not os.path.exists('assets'):
        os.mkdir('assets')

    r = requests.get(url)
    with open('assets/master.zip', 'wb') as f:
        f.write(r.content)

    extract_path = '.'
    with zipfile.ZipFile('assets/master.zip', 'r') as zip_ref:
        zip_ref.extractall(extract_path)

    if not os.path.exists('run.bat'):
        run_win = '''python batch_excel_email-master/main.py'''
        with open('run.bat', 'w') as run:
            run.writelines(run_win)

