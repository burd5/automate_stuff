import os
import pandas as pd
from datetime import date
from openpyxl import Workbook
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib as smtp
import config 
from Google import Create_Service

# Connect to Google Sheets API and get values from selected spreadsheet
FOLDER_PATH = r"C:\Users\Aburdett\Documents\PythonScripts\automatic_email"
CLIENT_SECRET_FILE = os.path.join(FOLDER_PATH, 'credentials.json')
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

g_sheet_id = '164uafkw8OyyYAa1xA_tstWmKzm4l6xW1V_jEVQaRn7c'
gsheet_name = 'Youth Violence Prevention Program Referral Form Responses'

service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)

gs = service.spreadsheets()
rows = gs.values().get(
    spreadsheetId=g_sheet_id,
    range='responses'
).execute()

# Create dataframe from rows and remove top row of index values from dataframe
df = pd.DataFrame(rows.get('values'))
df.columns = df.iloc[0]
df_new = df[1:]

# Save referral sheet to curent folder with name as {date}_referrals
curr_date = date.today()
output = "C://Users/Aburdett/Documents/PythonScripts/automatic_email/"
excel_file = df_new.to_excel(f'{output}{curr_date}_referrals.xlsx', index=False)

# Peek to see if everything looks correct
print(df_new)

# Email Information
msg = MIMEMultipart()
msg['From'] = 'aburdettgov@gmail.com'
msg['To'] = 'aburdettgov@gmail.com'
msg['Subject'] = 'Weekly Referral Form Sheet'

body = 'Hello, \n\nPlease find the attached Excel file with the updated referral responses.\n\nBest,\n\nAustin Burdette'
msg.attach(MIMEText(body, 'plain'))

# Attach file
filename = f'{curr_date}_referrals.xlsx'
attachment = open(filename, 'rb')

part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)

part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

msg.attach(part)

# Connect to server and send email, close connection
connection = smtp.SMTP_SSL('smtp.gmail.com', 465)
email_addr = 'aburdettgov@gmail.com'
email_passwd = config.PASSWORD
connection.login(email_addr, email_passwd)
connection.sendmail(email_addr, ['aburdett@auroragov.org', 'ssturge@justiceworksco.com'], msg.as_string())
connection.close()
