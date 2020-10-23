#!/usr/bin/env python3
#-------------------------------------------------------
import email, smtplib
import os.path

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
#-------------------------------------------------------

sender_email = "ecwhitney@icloud.com"
receiver_email = "erich@hbeng.com"
bcc_email = "ecwhitney@icloud.com"
smtp_address = "mail.myfairpoint.net"
filearg = "../21Aug2020.zip"

#-------------------------------------------------------

subject = "An email with attachment from Python"
body = "This is an email with attachment sent from Python\n\n"

#-------------------------------------------------------

# Create a multipart message and set headers
message = MIMEMultipart()
message["From"] = sender_email
message["To"] = receiver_email
message["Subject"] = subject
message["Bcc"] = bcc_email  # Recommended for mass emails

# Add body to email
message.attach(MIMEText(body, "plain"))

(dirname, filename) = os.path.split(filearg)

# Open PDF file in binary mode
with open("%s/%s" % (dirname, filename), "rb") as attachment:
    # Add file as application/octet-stream
    # Email client can usually download this automatically as attachment
    part = MIMEBase("application", "zip")
    part.set_payload(attachment.read())

# Encode file in ASCII characters to send by email    
encoders.encode_base64(part)

# Add header as key/value pair to attachment part
part.add_header(
    "Content-Disposition",
    "attachment; filename= %s" % filename,
)

# Add attachment to message and convert message to string
message.attach(part)
text = message.as_string()

try:	
   smtpObj = smtplib.SMTP(smtp_address)      
   smtpObj.sendmail(sender_email, receiver_email, text)        
   print("Successfully sent email")                    
except smtplib.SMTPException:                          
   print("Error: unable to send email")                
