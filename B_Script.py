#import statement if Required
from sys import *
from datetime import date
import pandas as pd
import xlrd
import smptlib
from email.message import EmailMessage
import schedule
import time

#Functions Used in Automation Script
def Sending_Mail(data):
  
  msg = EmailMessage();
  
  msg.set_content("Hello"+data[1]+",\n"+"Wishing You a Very Very Wonderful Birthday! May all your dreams comes true!!Have a great year ahead.💫️🤗️💥️🍫️💥️💥️")
  msg['Subject'] = "Happy Birthday(Wishes)";
  msg['From'] = 'abhigadhave97@gmail.com'
  msg['To'] = data[4]
  
  server = smtplib.SMTP_SSL("smtp.gmail.com",465)
	
	server.login("abhigadhave97@gmail.com","Your_Mail_Password")
	
	server.send_message(msg);
	
	server.quit()


def Receive_Data(strX):
  todayDate = date.today()
  
  stringD = (str)(todayDate)
  x = stringD.split("-")
  
  DOB = x[1]+"-"+x[2]
  
  workbook = pd.read_excel(strX)
  
  for i in range(0,len(workbook.index)):
    if(DOB == workbook['DOB'].iloc[i]):
      
      #To open Workbook
      wb = xlrd.open_workbook(strX)
      sheet = wb.sheet_by_index(0)
      Sending_Mail(sheet.row_values(i+1))
      del wb


#Entry Point of Automation Script
def main():
  
  print("--------------Automation Script : Abhijit Gadhave-----------------")
  print("Script Name : ",argv[0])
  print("Total Number of Arguments Accepted : ",len(argv))
  
  if(len(argv)!=2):
    print("Invalid Number of Arguments")
    exit()
  
  if(argv[1]=='-u' or argv[1]=='-U'):
    print("Usage : It is Used to Sending Mail to Happy Birthday Automatically With respective day!!!");
    exit();
  
  if(argv[1]=='-h' or argv[1]=='-H'):
    print("Help : Name_Of_Script  Argument")
    print("Argument : Name of that file where data can be stored");
    exit();
   
 try:
  schedule.every(1).minute.do(Receive_data(argv[1]))
  
  while(True):
    schedule.run_pending()
  
  except Exception as e:
    print("Exception : ",e);
  

#starter of the Automation Script
if __name__ == "__main__":
  main()
