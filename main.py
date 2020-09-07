import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib

#Sender's email credentials
SenderAddress = "Your Email"
password = "Your Password"

#Read's excel data
e = pd.read_excel("Contacts File")
c = pd.read_excel("Coupon Code File")
#Add the coupon code column to the email excel file (for simplicity)
e = e.drop(columns='contact_name')
e['couponcodes'] = c['Discount Code']
c = c[c['Status'] != 'Given']
#Create and send customzied email
for index, row in e.iterrows():

    #Setup Multi-part message
    msg = MIMEMultipart()
    msg['Subject'] = "Your Subject"
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
    fo = open(row['certificates'] + '.pdf', 'rb')
    attach = MIMEApplication(fo.read(), _subtype='pdf')
    fo.close()
    attach.add_header('Content-Disposition', 'attachment', filename='certificate.pdf')
    #attach pdf and the email body
    msg.attach(attach)
    msg.attach(HTML_Contents)
    msg['To'] = row['contact_email']
    msg['From'] = SenderAddress
    #send email
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        print('Mailing to: ' + row['contact_email'], row['couponcodes'])
        smtp.login(SenderAddress, password)
        smtp.send_message(msg)
        df = pd.read_excel("Coupon Code File", index_col='Discount Code')
        df.loc[row['couponcodes'], 'Status'] = 'Given'
        df.to_excel("Coupon Code File")
        

