import sys
import smtplib
import getpass
import xlrd
from random import seed
from random import randint
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

secretList = {}

def main():
    loc = '/Users/brycearcher/Documents/SecretSanta.xlsx'
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)

    for i in range(sheet.nrows):
        num = int(sheet.cell_value(i, 1))
        prov = str(sheet.cell_value(i, 2))
        picked = False
        secretList[sheet.cell_value(i, 0)] = [num, prov, picked]


    secret_santa()


def secret_santa():
    email = str(input("Email: "))
    pas = getpass.getpass(prompt="Password: ")

    person = ''
    personval =''

    try:
        smtp = "smtp.att.yahoo.com"
        port = 587

        server = smtplib.SMTP(smtp, port)
        server.starttls()
        server.login(email, pas)

        for k, v in secretList.items():
            rand = randint(0, len(secretList) - 1)
            person = list(secretList.keys())[rand]
            personval = list(secretList.values())[rand]
            while personval[2] is True or person is k:
                if len(list(secretList.keys())) <= 1:
                    break
                print('hi')
                rand = randint(0, len(secretList) - 1)
                person = list(secretList.keys())[rand]
                personval = list(secretList.values())[rand]

            personval[2] = True
            word = '' + str(v[0]) + '@'

            prov = v[1]
            if prov == 'att':
                word += 'mms.att.net'
            elif prov == 'verizon':
                word += 'vzwpix.com'
            # elif prov == 'tmobile':
            #     # word = '' + str(v[0]) + '@'
            #     word += 'tmomail.net'
            elif prov == 'sprint':
                word += 'pm.sprint.com'

            print(k)

            sms_gateway = word
            print(sms_gateway)

            msg = MIMEMultipart()
            msg['From'] = email
            msg['To'] = sms_gateway
            body = 'Hello ' + k + '!\n'
            body += 'This is a too long do not wanna read:\n You are getting a gift for ***' + person + \
                    '***. The rules are at the bottom. Please read those.\n\n'
            body += 'As we all know the holiday season for Thanksgiving and getting shit faced and ' \
                    'fucked up with your best friends has finally come to an end. However, ' \
                    'with the end of one holiday season comes the season of giving, receiving, ' \
                    'and getting ~f u c k e d u p~. With that being said, the time for Secret Santa ' \
                    'has come! Your going to be getting a gift for...' + person + '! Get them something they ' \
                    'will love, hate, cherish, frankly  I do not care (unless you got me, Bryce, then get me something ' \
                    'good). Good luck, have fun, and remember, let us all have fun and get fucked up when we can all ' \
                    'celebrate!\n\n'
            body += 'Again, your person is ***' + person + '***. Here are the rules:\n'
            body += '\t1.) Money range is $10-$20. Please stay in this cause we all broke college kids.\n'
            body += '\t2.) You must come to the agreed day of celebration. If you cannot make it, we can figure ' \
                    'something out where we can still get presents. That can be talked about.\n'
            body += '\t3.) Third and final rule, you must partake in the drinking and getting fucked up at the ' \
                    'celebration if there.\n'
            body += 'That is all. Let us have fun. Sorry fo the long message as I am on addy.'
            msg = MIMEText(body)
            # sms = msg.as_string()
            # server.sendmail(email, sms_gateway, sms)
            server.send_message(msg, email, sms_gateway)

        server.quit()

    except smtplib.SMTPServerDisconnected:
        print("Must input an email and password")
        sys.exit(0)


if __name__ == "__main__": main()