from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import pandas as pd
from datetime import date

today = date.today()

def initServer():
  server = smtplib.SMTP(host="smtp.gmail.com",port=587)
  server.ehlo()
  server.starttls()
  server.login("amrita.notification@gmail.com","mqqnqohxsefhxdse")
  return server

def initMessageHeader():
  message = MIMEMultipart()
  message["from"] = "Amrita Notification"
  messageBody = '''Respected Sir/Ma'am,

This is an automated mail from Amrita Notification regarding your Examination Duty on '''+today.strftime("%B %d, %Y")+'''
Have a Nice Day!

Regards
Amrita Notifications'''
  
  message.attach(MIMEText(messageBody))
  return message
  

if __name__=='__main__':
  print("Starting Amrita Notifications")
  df = pd.read_excel("/content/Book2.xlsx")
  server = initServer()
  for faculty in df['Email']:
    message = initMessageHeader()
    message["to"] = faculty
    message["subject"] = "Exam Duty Notification"
    server.send_message(message)
  server.close()
  print("Messages Sent Succesfully")
