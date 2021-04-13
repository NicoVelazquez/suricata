from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
import os


def read_sheet():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('suricata-mails-176e5e4e33a7.json', scope)
        client = gspread.authorize(credentials)
        sheet = client.open('Suricata Mails').get_worksheet(3)
        rows = sheet.get_all_records()

        for row in rows:
            if row['Email']:
                send_email(row['Name'], row['Location'], row['Email'])
            else:
                print(row['Company Name'] + ' has no email.')

    except Exception as e:
        print(e)


def send_email(c_name, c_location, c_email):
    gmail_user = os.environ['gmail_user']
    gmail_password = os.environ['gmail_password']

    message = MIMEMultipart("alternative")
    message["Subject"] = "Call this week?"
    message["From"] = "Nicolas Velazquez"
    message["To"] = c_email

    html = """\
    <html>
      <body>
        <p>Hello %s,
        <br>
        Just doing a follow-up!
        <br><br>
        I’m reaching out on behalf of Suricata and would really appreciate if we could meet up. 
        We believe that our product can really improve your customer support service.
        <br><br>
        By automating most of the most recurring clients questions, the your support to the clients would improve 
        drastically. 
        Not only would you decompress customer service but you would also show a really effective and fast attention 
        towards the clients.
        <br><br>
        You can check our website at http://suricata.la for further information or feel free to schedule a call 
        using the link below.
        <br>https://calendly.com/d/mmcw-2ksh/suricata
        <br><br>
        Kind regards,
        <br><br>
        </p>
        <h2 style="color: #666666; margin: 0">
            Nicolas Velazquez
        </h2>
        <p style="color: #666666; margin: 0">
            <span style="font-style: italic">Account Manager</span>
            <br>
            Mobile: +1 (611) 436-3814
            <br>
            WhatsApp: +54 9 11 5573-8585
            <br>
            Castillo 1366 - CABA - Argentina
            <br>
           <a href="http://www.suricata.la">www.suricata.la</a> 
        </p>
        <img src="cid:image1">
      </body>
    </html>
    """ % (c_name)

    part1 = MIMEText(html, "html")
    message.attach(part1)

    fp = open('./suricata-logo.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgImage.add_header('Content-ID', '<image1>')
    message.attach(msgImage)

    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        server.login(gmail_user, gmail_password)
        server.sendmail(gmail_user, c_email, message.as_string())
        server.close()

        print('Email sent to ' + c_email)
    except Exception as e:
        print('There was an error: ')
        print(e)


if __name__ == "__main__":
    read_sheet()
