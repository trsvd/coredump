from datetime import datetime
import imaplib
import email
# import html2text
from os import path

class MailClient(object):
    def __init__(self):
        self.m = imaplib.IMAP4_SSL('your.server.com')
        self.Login()


    def Login(self):
        result, data = self.m.login('login@domain.com', 'p@s$w0rd')
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

        message = email.message_from_string(data[0][1])
        res = {
            'From' : email.utils.parseaddr(message['From'])[1],
            'From name' : email.utils.parseaddr(message['From'])[0],
            'Time' : datetime.fromtimestamp(email.utils.mktime_tz(email.utils.parsedate_tz(message['Date']))),
            'To' : message['To'],
            'Subject' : email.Header.decode_header(message["Subject"])[0][0],
            'Text' : '',
            'File' : None
        }

        for part in message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get_content_maintype() == 'text':
                # reading as HTML (not plain text)

                _html = part.get_payload(decode = True)
                res['Text'] = html2text.html2text(_html)

            elif part.get_content_maintype() == 'application' and part.get_filename():
                fname = path.join("your/folder", part.get_filename())
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