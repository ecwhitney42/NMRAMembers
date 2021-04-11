#!/usr/bin/env python3
#-------------------------------------------------------
#import email, smtplib
#import os.path
#
#from email import encoders
#from email.mime.base import MIMEBase
#from email.mime.multipart import MIMEMultipart
#from email.mime.text import MIMEText
##-------------------------------------------------------
#
#sender_email = "ecwhitney@icloud.com"
#receiver_email = "erich@hbeng.com"
#bcc_email = "ecwhitney@icloud.com"
#smtp_address = "mail.myfairpoint.net"
#filearg = "../21Aug2020.zip"
#
##-------------------------------------------------------
#
#subject = "An email with attachment from Python"
#body = "This is an email with attachment sent from Python\n\n"
#
##-------------------------------------------------------
#
## Create a multipart message and set headers
#message = MIMEMultipart()
#message["From"] = sender_email
#message["To"] = receiver_email
#message["Subject"] = subject
#message["Bcc"] = bcc_email  # Recommended for mass emails
#
## Add body to email
#message.attach(MIMEText(body, "plain"))
#
#(dirname, filename) = os.path.split(filearg)
#
## Open PDF file in binary mode
#with open("%s/%s" % (dirname, filename), "rb") as attachment:
#    # Add file as application/octet-stream
#    # Email client can usually download this automatically as attachment
#    part = MIMEBase("application", "zip")
#    part.set_payload(attachment.read())
#
## Encode file in ASCII characters to send by email
#encoders.encode_base64(part)
#
## Add header as key/value pair to attachment part
#part.add_header(
#    "Content-Disposition",
#    "attachment; filename= %s" % filename,
#)
#
## Add attachment to message and convert message to string
#message.attach(part)
#text = message.as_string()
#
#try:
#   smtpObj = smtplib.SMTP(smtp_address)
#   smtpObj.sendmail(sender_email, receiver_email, text)
#   print("Successfully sent email")
#except smtplib.SMTPException:
#   print("Error: unable to send email")
#
   
   
#####################
# Python Send Email #
#####################

#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
##                                                                                                                                                                                                                                                                                                                       
## Imports                                                                                                                                                                                                                                                                                                               
##                                                                                                                                                                                                                                                                                                                       
#import os                                                                                                                                                                                                                                                                                                               
#import time                                                                                                                                                                                                                                                                                                             
#import ssl                                                                                                                                                                                                                                                                                                              
#import imaplib                                                                                                                                                                                                                                                                                                          
#import smtplib                                                                                                                                                                                                                                                                                                          
#import email                                                                                                                                                                                                                                                                                                            
#from email import encoders                                                                                                                                                                                                                                                                                              
#from email.mime.base import MIMEBase                                                                                                                                                                                                                                                                                    
#from email.mime.multipart import MIMEMultipart                                                                                                                                                                                                                                                                          
#from email.mime.text import MIMEText                                                                                                                                                                                                                                                                                    
#from email.headerregistry import Address                                                                                                                                                                                                                                                                                
#from email.message import EmailMessage                                                                                                                                                                                                                                                                                  
#                                                                                                                                                                                                                                                                                                                        
##                                                                                                                                                                                                                                                                                                                       
## Help Text                                                                                                                                                                                                                                                                                                             
##                                                                                                                                                                                                                                                                                                                       
#print("\n -----------------------------------------------\n =[           PYTHON E-MAIL SENDER            ]=\n -----------------------------------------------\n = Use \\ to break the line. Recipient format   =\n = can be email@srv.com or Name <mail@srv.com> =\n -----------------------------------------------\n")
#                                                                                                                                                                                                                                                                                                                        
#                                                                                                                                                                                                                                                                                                                        
##                                                                                                                                                                                                                                                                                                                       
## Create Message                                                                                                                                                                                                                                                                                                        
##                                                                                                                                                                                                                                                                                                                       
#def create_email_message(from_address, to_address, subject, body):                                                                                                                                                                                                                                                      
#    msg = EmailMessage()                                                                                                                                                                                                                                                                                                
#    msg['From'] = "ecwhitney@yahoo.com"                                                                                                                                                                                                                                                                                 
#    msg['To'] = "ecwhitney@icloud.com"                                                                                                                                                                                                                                                                                  
#    msg['Subject'] = subject                                                                                                                                                                                                                                                                                            
#    msg.set_content(body)                                                                                                                                                                                                                                                                                               
#    return msg                                                                                                                                                                                                                                                                                                          
#                                                                                                                                                                                                                                                                                                                        
##                                                                                                                                                                                                                                                                                                                       
## Message Data                                                                                                                                                                                                                                                                                                          
##                                                                                                                                                                                                                                                                                                                       
#if __name__ == '__main__':                                                                                                                                                                                                                                                                                              
#    msg = create_email_message(                                                                                                                                                                                                                                                                                         
#        from_address='ecwhitney@yahoo.com',                                                                                                                                                                                                                                                                             
#        to_address='ecwhitney@icloud.com', #input(' Recipient: '),                                                                                                                                                                                                                                                      
#        subject='test message', #input(' Subject: '),                                                                                                                                                                                                                                                                   
#        body='This is a test message') #input('\n Message: ').replace('\\', '\n'))                                                                                                                                                                                                                                      
#                                                                                                                                                                                                                                                                                                                        
##                                                                                                                                                                                                                                                                                                                       
## Account Details                                                                                                                                                                                                                                                                                                       
##                                                                                                                                                                                                                                                                                                                       
#login_email = 'ecwhitney@yahoo.com'                                                                                                                                                                                                                                                                                     
#login_passwd = '&V67cN-S:N77Y5i' #input('\n Password: ')                                                                                                                                                                                                                                                                
#                                                                                                                                                                                                                                                                                                                        
##                                                                                                                                                                                                                                                                                                                       
## Send Data to Server                                                                                                                                                                                                                                                                                                   
##                                                                                                                                                                                                                                                                                                                       
#try:                                                                                                                                                                                                                                                                                                                    
#    with smtplib.SMTP('smtp.mail.yahoo.com', port=587) as smtp_server:                                                                                                                                                                                                                                                  
##        smtp_server.ehlo()                                                                                                                                                                                                                                                                                             
#        smtp_server.starttls()                                                                                                                                                                                                                                                                                          
##        smtp_server.login(login_email, login_passwd)                                                                                                                                                                                                                                                                   
#        smtp_server.send_message(msg)                                                                                                                                                                                                                                                                                   
#                                                                                                                                                                                                                                                                                                                        
##                                                                                                                                                                                                                                                                                                                       
## Print Delivery Status                                                                                                                                                                                                                                                                                                 
##                                                                                                                                                                                                                                                                                                                       
#        print('\n \u2713 Email sent successfully.')                                                                                                                                                                                                                                                                     
#                                                                                                                                                                                                                                                                                                                        
#except Exception as e:                                                                                                                                                                                                                                                                                                  
#    print('\n Error: ',e, '\n Email not sent!')                                                                                                                                                                                                                                                                         
#                                                                                                                                                                                                                                                                                                                        
###                                                                                                                                                                                                                                                                                                                      
### Save Message to INBOX.Sent                                                                                                                                                                                                                                                                                           
###                                                                                                                                                                                                                                                                                                                      
##text = msg.as_string()                                                                                                                                                                                                                                                                                                 
##try:                                                                                                                                                                                                                                                                                                                   
##    with imaplib.IMAP4_SSL('imap.mail.me.com', 993) as imap_server:                                                                                                                                                                                                                                                    
##        imap_server.login(login_email, login_passwd)                                                                                                                                                                                                                                                                   
##        imap_server.append('INBOX.Sent', '\\Seen', imaplib.Time2Internaldate(time.time()), text.encode('utf8'))                                                                                                                                                                                                        
##        imap_server.logout()                                                                                                                                                                                                                                                                                           
##                                                                                                                                                                                                                                                                                                                       
###                                                                                                                                                                                                                                                                                                                      
### Print Saving Status                                                                                                                                                                                                                                                                                                  
###                                                                                                                                                                                                                                                                                                                      
##        print(' \u2713 Saved to INBOX.sent successfully.')                                                                                                                                                                                                                                                             
##except Exception as e:                                                                                                                                                                                                                                                                                                 
##    print('\n Error: ',e, '\n Could not save email to INBOX.sent!')                                                                                                                                                                                                                                                    
#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

import smtplib

sender = 'ecwhitney@icloud.com'
receivers = ['ecwhitney@icloud.com']

message = """From: Erich Whitney <ecwhitney@icloud.com>
To: Erich Whitney <ecwhitney@icloud.com>
Subject: SMTP e-mail test

This is a test e-mail message.
"""

try:
   smtpObj = smtplib.SMTP('localhost')
   smtpObj.sendmail(sender, receivers, message)
   print("Successfully sent email")
except Exception:
   print("Error: unable to send email")
   
   