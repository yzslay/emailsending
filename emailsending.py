import smtplib
from email.message import EmailMessage
from string import Template
from pathlib import Path
import pandas as pd

your_email = "[your email]"  # account
your_password = "[Password]"  # Password

# 用Path找html並讀取，接著用Template帶入來改關鍵字
html = Template(
    Path('D:/python3/pythonmail/mailtemplate.html').read_text(encoding="utf-8"))
emailcontent = EmailMessage()
emailcontent['subject'] = '請盡快還款'
emailcontent['from'] = 'Roger Hsiao'
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

email_list = pd.read_excel('D:/python3/pythonmail/Maillist.xlsx')  # 設定檔案位置
names = email_list['NAME']  # 讀取Name
emails = email_list['EMAIL']  # 讀取Email
company = email_list['COMPANY']  # 讀取COMPANY


for i in range(len(emails)):

    # for every record get the name and the email addresses
    name = names[i]
    email = emails[i]
    company = company[i]
    # the message to be emailed 更換
    emailcontent.set_content(html.substitute(
        {'name': name, 'company': company}), 'html')
    # sending the email
    print(emailcontent)
    server.sendmail(your_email, email, emailcontent.as_string())

server.close()
