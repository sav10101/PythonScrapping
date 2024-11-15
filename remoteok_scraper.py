import requests
import xlwt
from xlwt import Workbook
import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate

BASE_URL = 'https://remoteok.com/api/'
USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36'
REQUEST_HEADER = {
    'User-Agent': USER_AGENT,
    'Accept-Language': 'en-US, en;q=0.5',
}

SMTP_SERVER = 'smtp.outlook.com'
SMTP_PORT = 587
SMTP_USER = ''
SMTP_PASSWORD = ''
TO_EMAIL = ''
FROM_EMAIL = ''
SUBJECT = 'Job Postings Report'

def get_job_postings():
    res = requests.get(url=BASE_URL, headers=REQUEST_HEADER)
    return res.json()

def output_jobs_to_xls(data):
    wb = Workbook()
    job_sheet = wb.add_sheet('Jobs')
    
    headers = list(data[0].keys())
    for i, header in enumerate(headers):
        job_sheet.write(0, i, header)
    
    for row_num, job in enumerate(data, start=1):
        for col_num, header in enumerate(headers):
            job_sheet.write(row_num, col_num, job.get(header, ''))
    
    filename = 'jobs_report.xls'
    wb.save(filename)
    return filename

def send_email_with_attachment(subject, body, to_email, from_email, smtp_server, smtp_port, smtp_user, smtp_password, attachment_path):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Date'] = formatdate(localtime=True)

    msg.attach(MIMEText(body))

    with open(attachment_path, 'rb') as f:
        part = MIMEApplication(f.read(), Name=basename(attachment_path))
    part['Content-Disposition'] = f'attachment; filename="{basename(attachment_path)}"'
    msg.attach(part)

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(from_email, to_email, msg.as_string())

if __name__ == "__main__":
    job_data = get_job_postings()
    filename = output_jobs_to_xls(job_data)
    
    send_email_with_attachment(
        SUBJECT,
        'Informe de trabajos remotos.',
        TO_EMAIL,
        FROM_EMAIL,
        SMTP_SERVER,
        SMTP_PORT,
        SMTP_USER,
        SMTP_PASSWORD,
        filename
    )