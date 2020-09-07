import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib

#Sender's email credentials
SenderAddress = "Email"
password = "Password"

#Read's excel data
e = pd.read_excel("~/Downloads/contacts.xlsx")
c = pd.read_excel("~/Downloads/coupons.xlsx")
#Add the coupon code column to the email excel file (for simplicity)
e = e.drop(columns='contact_name')
e = e.replace('mihir_certificate', '/Users/arizsiddiqui/Downloads/mihir_certificate.pdf')
e['couponcodes'] = c['Discount Code']
#Create and send customzied email
for index, row in e.iterrows():

    #Setup Multi-part message
    msg = MIMEMultipart()
    msg['Subject'] = "Test Mails"
    #Setup Message (Customize as required)
    message= """\
<html>
  <body>
  <!-- Sample HTML, Customize to suit your needs -->
    <h1> Hey! Your Coupon code is: {} </h1>
  </body>
</html>
""".format(row['couponcodes'])
    #Convert the html string to the actual body
    HTML_Contents = MIMEText(message, 'html')
    #open and initialize the pdf file to attach
    fo = open(row['certificates'], 'rb')
    attach = MIMEApplication(fo.read(), _subtype='pdf')
    fo.close()
    attach.add_header('Content-Disposition', 'attachment', filename='certificate.pdf')
    #attach pdf and the email body
    msg.attach(attach)
    msg.attach(HTML_Contents)
    msg['To'] = row['contact_email']
    #send email
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        print('Mailing to: ' + row['contact_email'], row['couponcodes'])
        smtp.login(SenderAddress, password)
        #print(msg)
        smtp.send_message(msg.as_string(), SenderAddress, )
