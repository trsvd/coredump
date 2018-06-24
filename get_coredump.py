import quopri

import imaplib
import email
imaplib.IMAP4.debug = imaplib.IMAP4_SSL.debug = 1

username,passwd = ('hfyv','AZ_173_0()')

def view_emails():
    con = imaplib.IMAP4_SSL('imap.yandex.ru',993)
    con.login(username, passwd)
    con.select()
    typ, data = con.search(None, '(UNSEEN)')
    c = 0
    for num in data[0].split():
        typ, data = con.fetch(num, '(RFC822)')
        c +=1
        text = data[0][1]
        try:
            msg = email.message_from_string(data[0][1])
        except TypeError:
            msg = email.message_from_bytes(data[0][1])

        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            filename = quopri.decodestring(part.get_filename())

            data = part.get_payload(decode=True)
            if not data:
                continue
            filename = 'd:/Temp/attachments/' + filename
            f  = open(filename, 'w')
            f.write(data)
            f.close()

    con.close()
    con.logout()



# def SaveLstToFile(filename, lst):
#     file = open(filename, 'w')
#
#     string_value = ''
#
#     if type(lst) is list:
#         for value in lst:
#             string_value = value
#             file.write(string_value + '\n')
#
#     else:
#         for key, value in lst.iteritems():
#             if type(value) is list:
#                 string_value = key + ': ' + '\t'.join(value)
#             elif type(value) is dict:
#                 string_value = key + ': ' + '\n'
#                 for sub_key, sub_value in value.iteritems():
#                     if type(sub_value) is list:
#                         string_value = string_value + '\t' + sub_key + ': ' + '\t'.join(sub_value)
#                     else:
#                         string_value = string_value + '\t' + sub_key + ': ' + sub_value
#                     string_value = string_value + '\n'
#             else:
#                 string_value = key + ': ' + value
#
#             file.write(string_value + '\n')
