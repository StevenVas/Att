import os
import smtplib
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

class sendEmail(object):
    def __init__(self,username,password,receiver,title,content,file=None,ssl=False,email_host='smtp.163.com',port=25,ssl_port=465):
        self.username=username
        self.password=password
        self.receiver=receiver
        self.title=title
        self.content=content
        self.file=file
        self.ssl=ssl
        self.email_host=email_host
        self.port=port
        self.ssl_port=ssl_port

    def send_email(self):
        msg = MIMEMultipart()
        #发送内容对象
        if self.file:
            file_name=os.path.split(self.file)[-1] #只取文件名，不取文件路径
            try:
                f=open(self.file,'rb').read()
            except Exception as e :
                raise Exception('附件打不开')
            else:
                att =MIMEText(f,"base64","utf-8")
                att["Content-Type"]='application/octet-stream'
                #处理文件名为中文名
                new_file_name= '=?utf-8?b?' +base64.b64encode(file_name.encode()).decode()+'?='
                # att["Content-Disposition"]='attachment,filename="%s"'%(new_file_name)
                att.add_header('Content-Disposition', 'attachment', filename='Phoenix平台项目群_SIT测试案例与计划（合稿）_V0.2_20201106.xlsx')
                msg.attach(att)
        msg.attach(MIMEText(self.content)) #邮件正文内容
        msg['subject'] = self.title #邮件主题
        msg['From'] = self.username #发送者账号
        msg['To'] =','.join(self.receiver)#接收者账号列表
        if self.ssl:
            self.smtp =smtplib.SMTP_SSL(self.email_host,port=self.ssl_port)
        else:
            self.smtp =smtplib.SMTP_SSL(self.email_host,port=self.port)
        #发送邮件服务器的对象
        self.smtp.login(self.username,self.password)
        try:
            self.smtp.sendmail(self.username,self.receiver,msg.as_string())
            pass
        except Exception as e:
            print('出错了。。',e)
        else:
            print('发送Success,So Sexy!')
        self.smtp.quit()

if __name__ == '__main__':
    m = sendEmail(
        username='Wanyaze@163.com',
        password='RQJXBCPAZNNHYMML',
        receiver=['Wanyaze@163.com'],
        title='WangHUI So Sexy',
        content='接口自动化测试报告',
        file=r'C:\Users\16156\Desktop\接口测试\Phoenix平台项目群_SIT测试案例与计划（合稿）_V0.2_20201106.xlsx',
        ssl=True,

    )
m.send_email()


