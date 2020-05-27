import smtplib
from email.message import EmailMessage
from email.headerregistry import Address, Group
import email.policy
import mimetypes
import base64


class SendEmail(object):
    """
    python3 发送邮件类
    格式： html
    可发送多个附件
    """
    def __init__(self, smtp_server, smtp_user, smtp_passwd, sender, recipient):
        # 发送邮件服务器,常用smtp.163.com
        self.smtp_server = smtp_server
        # 发送邮件的账号
        self.smtp_user = smtp_user
        # 发送账号的客户端授权码
        self.smtp_passwd = smtp_passwd
        # 发件人
        self.sender = sender
        # 收件人
        self.recipient = recipient

        # Use utf-8 encoding for headers. SMTP servers must support the SMTPUTF8 extension
        # https://docs.python.org/3.6/library/email.policy.html
        self.msg = EmailMessage(email.policy.SMTPUTF8)

    def set_header_content(self, subject, content):
        """
        设置邮件头和内容
        :param subject: 邮件题头
        :param content: 邮件内容
        :return:
        """
        self.msg['From'] = self.sender
        self.msg['To'] = self.recipient
        self.msg['Subject'] = subject
        self.msg.set_content(content, subtype="html")

    def set_accessories(self, path_list: list):
        """
        添加附件
        :param path_list: [{"path": ""}, {"name": ""}]
        :return:
        """
        for path_dict in path_list:
            filename = path_dict['path']
            name = path_dict['name']
            ctype, encoding = mimetypes.guess_type(filename)
            if ctype is None or encoding is not None:
                # No guess could be made, or the file is encoded (compressed), so
                # use a generic bag-of-bits type.
                ctype = 'application/octet-stream'
            maintype, subtype = ctype.split('/', 1)
            with open(filename, 'rb') as fp:
                self.msg.add_attachment(fp.read(), maintype, subtype, filename=self.dd_b64(name))

    def send_email(self):
        """
        发送邮件
        :return:
        """
        with smtplib.SMTP_SSL(self.smtp_server) as smtp:
            # HELO向服务器标志用户身份
            smtp.ehlo_or_helo_if_needed()
            # 登录邮箱服务器
            smtp.login(self.smtp_user, self.smtp_passwd)
            print("Email:{}==>{}".format(self.sender, self.recipient))
            smtp.send_message(self.msg)
            print("Sent succeed!")

    @staticmethod
    def dd_b64(param):
        """
        对邮件header及附件的文件名进行两次base64编码，防止outlook中乱码。
        email库源码中先对邮件进行一次base64解码然后组装邮件
        :param param: 需要防止乱码的参数
        :return:
        """
        param = '=?utf-8?b?' + base64.b64encode(param.encode('UTF-8')).decode() + '?='
        param = '=?utf-8?b?' + base64.b64encode(param.encode('UTF-8')).decode() + '?='
        return param


if __name__ == '__main__':
    smtp_server = "smtp.163.com"
    smtp_user = "a.163.com"  # 发送邮件的账号
    smtp_passwd = "aaaaaa"  # 发送账号的客户端授权码
    # 直接用字符串也行，这里为了方便使用 display name, 相当于 "测试-发送者 <a@163.com>"
    # 发送号码要与登录号码一致
    sender = Address("测试-发送者", "a", "163.com")
    # 发送多个人
    recipient = Group(addresses=(Address("测试-接受者", "b", "qq.com"), Address("测试-接受者", "c", "qq.com")))
    subject = "测试专用"
    content = """
        <html>
            <h2 style='color:red'>将进酒</h2>
            <p>人生得意须尽欢，莫使金樽空对月。</p>
            <p>天生我材必有用，千金散尽还复来。</p>
        </html>
        """
    path_list = [{"path": r"./测试文件123abc.txt",
                  "name":  "测试文件123abc.txt"}]
    sd = SendEmail(smtp_server, smtp_user, smtp_passwd, sender, recipient)  # 创建对象
    sd.set_header_content(subject, content)  # 设置题头和内容
    sd.set_accessories(path_list)  # 添加附件
    sd.send_email()  # 发送邮件

