import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib

#Sender's email credentials
SenderAddress = "<Email Address of sender>"
password = "password of sender"

#Read's excel data
e = pd.read_excel("Email.xlsx")
c = pd.read_excel("Coupon Codes.xlsx")
#Add the coupon code column to the email excel file (for simplicity)
e['couponcodes'] = c['Discount Code]
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(SenderAddress, password)

#Setup Multi-part message
msg = MIMEMultipart('alternative')
msg['Subject'] = ""

#Create and send customzied email
for email, filename, cc in e.iterrows():

    #Setup Message (Customize as required)
    message= """\
<html>
  <body>
  <!-- Sample HTML, Customize to suit your needs -->
    <p>Hi,<br>
       How are you?<br>
       <a href="http://www.realpython.com">Real Python</a>
       has many great tutorials.
    </p>
  </body>
</html>
"""
    #Convert the html string to the actual body
    HTML_Contents = MIMEText(message, 'html')
    #open and initialize the pdf file to attach
    fo = open(filename, 'rb')
    attach = MIMEApplication(fo.read(), _subtype='pdf')
    fo.close()
    attach.add_header('Content-Disposition', 'attachment', filename=filename)
    #attach pdf and the email body
    msg.attach(attach)
    msg.attach(HTML_Contents)
    #send email
    s.sendmail(SenderAddress, email, msg.as_string())
server.quit()
