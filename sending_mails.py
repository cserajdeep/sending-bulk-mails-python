import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import re
import config

def send_email(subject, msg, to_):
    global mail_count
    
    #Setup the MIME
    message = MIMEMultipart("alternative")
    message['From'] = config.EMAIL_ADDRESS
    message['To'] = to_
    message['Subject'] = subject
    # attach the body with the msg instance 
    message.attach(MIMEText(msg, 'html')) 
    msg = message.as_string() 
    try:
      server = smtplib.SMTP('smtp.gmail.com:587')
      server.starttls()
      server.login(config.EMAIL_ADDRESS, config.PASSWORD)
      server.sendmail(config.EMAIL_ADDRESS, to_, msg)
      server.quit()
      
      mail_count = mail_count + 1
      print(mail_count,". Success: Email sent to ", to_)
    except:
      print("Email failed to send.")
      print("Total", mail_count," Mail Sent!!")
      

###---------------------------MAIN--------------------------------------
# Here, the given .xlsx file name is participants. stored at the same folder
data = pd.read_excel('YOUR_MAILLIST.xlsx')

name_list = data["Name"].tolist()    #NAME HEADER
mailid_list = data["MailId"].tolist()   #MAIL-ID HEADER
email_pattern= re.compile("^.+@.+\..+$")

mail_count = 0
for i,mail in enumerate(mailid_list):
  author = name_list[i]
  fname = author.split(' ')[0]
  
  #-----------------------------------------------------------------------------
  # CUSTOM MESSAGE (HTML/PLAIN)   
  msg = '''Dear '''+ author +''',<br><br>
  I hope that this email finds you well.<br><br>
  If you have already received this message, please ignore it. <br><br>
  Thanks & Regards,<br>
  NAME OF THE SENDER<br>
  AFFILIATION OF THE SENDER<br>
  '''

  subject = "SUBJECT OF THE MAIL"
  try:
    if( email_pattern.search(mail)):
      send_email(subject, alumni_msg, mail)
  except:
    print("Mail id is not correct.", mail)

print("Total", mail_count," Mail Sent!!")
