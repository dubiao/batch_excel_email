
TEMPLATE_TITLE_MAP = {
    'seq': '序号',
    'name': '姓名',
    'email': '邮箱',
    # 'id_number': '身份证号码',
    'bank_account': '银行卡号',
    'salary': '基本薪资',
    'leave_money': '考勤扣款-事假',
    'sick_money': '考勤扣款-病假',
    'no_job_money': '考勤扣款-入/离职缺勤',
    'additions': ["通讯补贴", "油费补贴", "交通补贴", "全勤奖", "餐补", "过节费", "工作餐补", "误餐补助"],
    'salary_fix': "补发薪资",
    'salary_should': "应发工资",
    'no_concat_titles': ['五险一金专项扣除-个人部分'],
    'baoxian_moneys': ['养老', '医疗', '失业'],
    'gongjijin_money': "公积金",
    'salary_before_tax': "税前工资",
    'tax_money': "薪资个税",
    'salary_after_tax': "税后工资",
    'money_this_month': "本月实发金额",
    'out_day': "离职日期"
}
TABLE_VIEW_HEAD = [
    'seq',
    'name',
    'email',
    'out_day',
    'salary',
    'salary_should',
    'salary_before_tax',
    'tax_money',
    'money_this_month'
]
EMAIL_CONTENT = '''
<html><body>
  <p>亲爱的 {name}：</p>
  <p>附件中是您{month}的工资信息单，请查收。如有问题，请及时回复本邮件。</p>
  <hr/>
  <p>财务部</p>
  <p>{today}</p>
</body></html>
'''
EMAIL_SUBJECT = '【工资条】{name} - {month}'
DEFAULT_CONF = {
    'file_path': '.',
    'file_sheet_name_fmt': '工资表{month_code}',
    'template_title_map': TEMPLATE_TITLE_MAP,
    'table_view_head': TABLE_VIEW_HEAD,
    'generate_file': {
        'email_content': EMAIL_CONTENT,
        'email_subject': EMAIL_SUBJECT,
        'template_file_path': 'assets/salary_email_template.docx',
        'save_file_path': 'assets/files/{today}',
        'save_file_name': '{name}-{month}-工资',
        'convert_to_pdf': True,  # Only works on Windows with Office installed
    },
    'tester_email': None,
}
