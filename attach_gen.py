from docxtpl import DocxTemplate
import os
import time

try:
    from comtypes import client
except ImportError:
    client = None


class EmailGenerator:
    def __init__(self, generate_config, year, month, target_folder=None):
        if not os.path.exists(generate_config['template_file_path']):
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
            if self.convert_pdf(to_file, to_pdf_file):
                return to_pdf_file
        return to_file

    def convert_pdf(self, infile, outfile):
        if not client:
            return False
        format_pdf = 17
        word = client.CreateObject('Word.Application')
        doc = word.Documents.Open(infile)
        doc.SaveAs(outfile, FileFormat=format_pdf)
        doc.Close()
        word.Quit()
        return True

    def make_email(self, params):
        if params is None:
            print('error params')
            return None
        subject = self.email_content.format(
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
