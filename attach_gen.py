from docxtpl import DocxTemplate
import os
import time

try:
    from win32com import client
except ImportError:
    client = None


class EmailGenerator:
    def __init__(self, generate_config, year, month, target_folder=None):
        if not os.path.exists(generate_config['template_file_path']):
            print(generate_config['template_file_path'])
            raise IOError('Template file not found')
        self.template_path = generate_config['template_file_path']
        self.for_month = '%d年%02d月' % (year, month)
        self.for_month_code = '%d%02d' % (year, month)
        if target_folder is None:
            target_folder = generate_config['save_file_path'].format(today=time.strftime("%Y-%m-%d", time.localtime()))
        if not os.path.exists(target_folder):
            os.makedirs(target_folder)
        self.pdf = True
        self.target_path = '%s/%s.%s' % (target_folder, generate_config['save_file_name'], 'docx')
        self.email_content = generate_config['email_content']
        self.email_subject = generate_config['email_subject']
        self.generate_config = generate_config
        self.word = None

    def make_file(self, params, to_file=None, make=True):
        if params is None:
            print('error params')
            return None
        if to_file is None:
            to_file = self.target_path.format(name=params['name'], month=self.for_month)
        to_pdf_file = '%s.pdf' % os.path.splitext(to_file)[0]
        if not make:
            return to_pdf_file if self.pdf else to_file
        tpl = DocxTemplate(self.template_path)
        tpl.render(params)
        tpl.render({'month': self.for_month})
        tpl.save(to_file)
        if self.pdf:
            if not self.word:
                try:
                    self.word = client.Dispatch('Word.Application')
                except Exception:
                    self.pdf = False
            if self.pdf and self.word:
                result_path = self.convert_pdf(os.path.abspath(to_file), os.path.abspath(to_pdf_file))
                if result_path:
                    return result_path
        return to_file

    def convert_pdf(self, in_file, out_file):
        if client is None:
            return False
        format_pdf = 17

        if not self.word:
            # noinspection PyBroadException
            try:
                self.word = client.Dispatch('Word.Application')
            except Exception:
                self.pdf = False
                pass
        if not self.word:
            return False
        # noinspection PyBroadException
        try:
            doc = self.word.Documents.Open(in_file)
            doc.SaveAs(out_file, FileFormat=format_pdf)
            doc.Close()
            # no quit
            # self.word.Quit()
            return out_file
        except AttributeError:
            self.pdf = False
            return False
        except Exception:
            return False

    def make_email(self, params):
        if params is None:
            print('error params')
            return None
        subject = self.email_subject.format(
            month=self.for_month,
            today=time.strftime("%Y-%m-%d", time.localtime()),
            **params
        )
        html = self.email_content.format(
            month=self.for_month,
            today=time.strftime("%Y-%m-%d", time.localtime()),
            **params
        )
        return {'subject': subject, 'content': html}
