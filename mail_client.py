from datetime import datetime
import imaplib
import email
import html2text
from os import path

import quopri
import base64

imaplib.IMAP4.debug = imaplib.IMAP4_SSL.debug = 1

class MailClient(object):
    def __init__(self):
        self.m = imaplib.IMAP4_SSL('imap.yandex.ru', 993)
        self.Login()


    def Login(self):
        result, data = self.m.login('hfyv', 'AZ_173_0()')
        if result != 'OK':
            raise Exception("Error connecting to mailbox: {}".format(data))


    def ReadLatest(self, delete = True):
        result, data = self.m.select("inbox")
        if result != 'OK':
            raise Exception("Error reading inbox: {}".format(data))
        if data == ['0']:
            return None
        latest = data[0].split()[-1]
        result, data = self.m.fetch(latest, "(RFC822)")
        if result != 'OK':
            raise Exception("Error reading email: {}".format(data))
        if delete:
            self.m.store(latest, '+FLAGS', '\\Deleted')

        try:
            message = email.message_from_string(data[0][1])
        except TypeError:
            message = email.message_from_bytes(data[0][1])

        res = {
            'From' : email.utils.parseaddr(message['From'])[1],
            'From name' : email.utils.parseaddr(message['From'])[0],
            'Time' : datetime.fromtimestamp(email.utils.mktime_tz(email.utils.parsedate_tz(message['Date']))),
            'To' : message['To'],
            # 'Subject' : email.Header.decode_header(message["Subject"])[0][0],
            'Subject' : base64.b64decode(message["Subject"]),
            'Text' : '',
            'File' : None
        }

        # print (quopri.decodestring(res['Subject'].encode('utf-8')).decode('utf-8'))

        for part in message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get_content_maintype() == 'text':
                # reading as HTML (not plain text)

                # _html = quopri.decodestring(part.get_payload(decode = True))
                _html = part.get_payload(decode = True)
                res['Text'] = html2text.html2text(_html)

            elif part.get_content_maintype() == 'application' and part.get_filename():
                fname = path.join("d:/Temp/attachments/", quopri.decodestring(part.get_filename()))
                attachment = open(fname, 'wb')
                attachment.write(part.get_payload(decode = True))
                attachment.close()
                if res['File']:
                    res['File'].append(fname)
                else:
                    res['File'] = [fname]

        return res


    def __del__(self):
        self.m.close()