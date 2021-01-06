# coding ;utf-8

import psycopg2
import smtplib
import os
import openpyxl
import datetime
# from impala.dbapi import connect
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage


# C:/Users/KZF/Desktop/test1/result/result_2.xlsx
# C:/Users/lenovo/Desktop/数据分析协助/result/result_2.xlsx
def report_send_email(filename, receiver, from_mail, from_mail_pass):
    receiver = '1369157509@qq.com'  # 收件人邮箱地址

    date_str = datetime.datetime.strftime(datetime.date.today() - datetime.timedelta(days=1), '%m%d')

    mail_txt = """
        附件是运营日报，请查收
    """
    msgRoot = MIMEMultipart('mixed')
    msgRoot['Subject'] = str(u'日报-%s' % date_str)
    msgRoot['From'] = '15172054680@163.com'
    msgRoot['To'] = receiver
    msgRoot["Accept-Language"] = "zh-CN"
    msgRoot["Accept-Charset"] = "ISO-8859-1,utf-8"

    msg = MIMEText(mail_txt, 'plain', 'utf-8')
    msgRoot.attach(msg)
    att = MIMEText(open('C:/Users/KZF/Desktop/test1/result/result_2.xlsx', 'rb').read(), 'base64', 'utf-8')

    # att = MIMEText(open(filename, 'rb').read(), 'base64', 'utf-8')
    att["Content-Type"] = 'application/octet-stream'
    # att["Content-Disposition"] = 'attachment; filename="日报2020%s.xlsx"'% date_str
    att.add_header('Content-Disposition', 'attachment', filename="日报2020%s.xlsx" % date_str)
    msgRoot.attach(att)

    mail_server = 'smtp.163.com'
    smtp = smtplib.SMTP(host=mail_server)
    print(smtp)
    smtp.login(from_mail, from_mail_pass)
    for k in receiver.split(','):
        smtp.sendmail('15172054680@163.com', k, msgRoot.as_string())
    smtp.quit()


# if __name__ == '__main__':
#
#     receiver = '1369157509@qq.com'  # 收件人邮箱地址
#     date_str = datetime.datetime.strftime(datetime.date.today(), '%m%d')
# 
#     mail_txt = """
#         附件是运营日报，请查收
#     """
#     msgRoot = MIMEMultipart('mixed')
#     msgRoot['Subject'] = str(u'日报-%s' % date_str)
#     msgRoot['From'] = 'ye_zhan_bo@163.com'
#     msgRoot['To'] = receiver
#     msgRoot["Accept-Language"] = "zh-CN"
#     msgRoot["Accept-Charset"] = "ISO-8859-1,utf-8"
#
#     msg = MIMEText(mail_txt, 'plain', 'utf-8')
#     msgRoot.attach(msg)
#
#     att = MIMEText(open('C:/Users/KZF/Desktop/test1/result/result_2.xlsx', 'rb').read(), 'base64', 'utf-8')
#     att["Content-Type"] = 'application/octet-stream'
#     # att["Content-Disposition"] = 'attachment; filename="日报2020%s.xlsx"'% date_str
#     att.add_header('Content-Disposition', 'attachment', filename="日报2020%s.xlsx" % date_str)
#     msgRoot.attach(att)
#     # https://www.cnblogs.com/xiaodai12138/p/10483158.html
#     mail_server = 'smtp.163.com'
#     smtp = smtplib.SMTP(host =mail_server)
#     smtp.login('15172054680@163.com', 'BTHSVBTDJRABAUOJ')
#     for k in receiver.split(','):
#         smtp.sendmail('15172054680@163.com', k, msgRoot.as_string())
#     smtp.quit()
