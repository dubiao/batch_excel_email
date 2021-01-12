# 基于Excel文档的内容发批量发送邮件
常用的场景是每月发工资条。

会持续抽象为可以基于Excel文件发送任何业务。

### 使用方法

1. 运行程序会根据config_DEFAULT 生成默认的配置文件。
  可以根据需要修改 config_DEFAULT 或 conf.json
2. 目前 main.py 是与业务偶合的。可以根据需要来修改菜单和其它逻辑
3. 目前邮箱相关的配置直接写在代码里。主要是考虑写到配置文件会导致密码外泄
4. 文件转为PDF的功能只在Windows上已经安装Office套件时可用
5. 每次在正式批量发送之前用菜单的前几项测试好，以防因程序bug发错而导致严重后果


### 配置说明

大部分配置的含义一望而知

1. 配置里的变量：
  1. {month} 操作的相应月,在main里 input_month 需要输入 202012 这样的日期，month代表 “2020年12月”
  2. {month_code} 如上输入的年月 202012
  3. {today} 当前的日期
  3. {name} Excel中读取的到每行 name 列的值
  4. {email} Excel中读取的到每行 email 列的值
2. template_title_map 是一个map，标识Excel中表头对应到程序里的变量。
  1. additions 的意思是一系列额外的项目，在生成文件时将只显示内容 > 0 的项目，项目名称 依次替换模板文件中的 money{0}，金额值替换 money{0}amount 。 {0} 从 2 开始递增。最大为4
  2. out_day 是与工资业务偶合的。标识一个已经离职的人，用于提示是否对此人发邮件。之后会抽象到 main 或 conf 里面方便其它业务使用
  3. 其它字段均直接替换模板文件中的变量标识
2. file_sheet_name_fmt 是指定默认的Sheet表名称格式
  * 如果找不到相应的表格，程序也会提示选择正确的数据表格
3. table_view_head 程序在控制台打印出的缩略表格需要显示的列名
4. generate_file 是生成文件或邮件相关的控制变量组
  1. convert_to_pdf 文件转为PDF的功能  只在Windows上已经安装Office套件时可用 需要使用 comtypes 模块
5. tester_email 测试人的邮箱地址。在使用main中的测试发送时 可以输入一个邮箱地址，或直接使用此地址。自动记录最后一次输入。

