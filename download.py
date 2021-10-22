import requests
import credentials
import csv
import openpyxl
import smtplib
from email.message import EmailMessage
from email.mime.base import MIMEBase
from email import encoders


def sendmailneu(FROM_EMAIL, TO_EMAIL, SUBJECT, BODY, ATTACHMENTS_LIST):
    msg = EmailMessage()
    msg["From"] = FROM_EMAIL
    msg["Subject"] = SUBJECT
    msg["To"] = TO_EMAIL
    msg.set_content(
        "Hier sollte ein Anschreiben folgen. Tut es das nicht, dann bitte ein ordentliches Mailprogramm verwenden!")
    msg.add_alternative(BODY, subtype='html')
    for i in range(0, len(ATTACHMENTS_LIST)):
        part = MIMEBase('application', "octet-stream")
        with open(ATTACHMENTS_LIST[i], 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        f'attachment; filename="{ATTACHMENTS_LIST[i]}"')
        msg.attach(part)
    s.send_message(msg)


def returnBody(ABSENDER):

    BODY = f"\
        <html>\
            <body>\
                <p>\
                    Es gibt neue Anmeldungen für den Tag der offenen Schule.\
                </p>\
                Freundliche Grüße,<br>\
                {ABSENDER}\
            </body>\
        </html>"
    return BODY


def csv_to_excel(csv_file, excel_file):
    csv_data = []
    with open(csv_file) as file_obj:
        reader = csv.reader(file_obj)
        for row in reader:
            csv_data.append(row)
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    for row in csv_data:
        sheet.append(row)
    workbook.save(excel_file)
    row_count = sheet.max_row
    return row_count


response = requests.get(credentials.NEXTCLOUDLINK,
                        auth=(credentials.USERNAME_FORM, credentials.PASSWORD_FORM))

with open("anmeldungen.csv", "w", newline="") as f:
    f.write(response.text)

anzahlAnmeldungen = csv_to_excel("anmeldungen.csv", "anmeldungen.xlsx") - 1
gespeichert = 0


with open("antworten.txt", "r") as f:
    gespeichert = f.readlines()[0]


if gespeichert != str(anzahlAnmeldungen):
    print("gespeichert: ", gespeichert)
    print("Anzahl Antworten im download: ", anzahlAnmeldungen)

    with open("antworten.txt", "w") as f:
        f.writelines(str(anzahlAnmeldungen))

    with open("anmeldungen.csv", "w", newline="") as f:
        f.write(response.text)

    ATTACHMENTS_LIST = [
        "anmeldungen.xlsx"
    ]

    s = smtplib.SMTP_SSL(credentials.SMTPSERVER)
    s.login(credentials.USERNAME_MAIL, credentials.PASSWORD_MAIL)
    sendmailneu(credentials.USERNAME_MAIL, credentials.RECIPIENT,
                credentials.SUBJECT, returnBody(credentials.ABSENDER), ATTACHMENTS_LIST)
    s.quit()
    print("Mail gesendet.")
else:
    print("keine neuen Einträge.")
