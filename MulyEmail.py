# Python code to send email to a list of 
# emails from a spreadsheet 
  
# import the required libraries 
import pandas as pd 
import smtplib 
  
# email and  password
your_email = "your email"
your_password = "password"
  
#Login in 
server = smtplib.SMTP_SSL('smtp.gmail.com', 465) 
server.ehlo() 
server.login(your_email, your_password) 
  
#Read the Exel
email_list = pd.read_excel('Name of the Exel.xls') 
  
#making 2 list with the names and the emails 
names = email_list['NAME'] 
emails = email_list['EMAIL'] 
  
 
for i in range(len(emails)): 
  
    name = names[i] 
    email = emails[i] 
    print(emails[i])
    
    mail_subject = 'subject' #Subject
    mail_message_body = 'message body' #message body
    message = f'''\
From: {your_email}
To: {email}
Subject: {mail_subject}
{mail_message_body}
''' 
    print(message)
     
    server.sendmail(your_email, [email], message) 

#Close smtp server 
server.close() 
