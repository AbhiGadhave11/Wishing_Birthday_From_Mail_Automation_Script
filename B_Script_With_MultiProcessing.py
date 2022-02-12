from sys import *
from datetime import date
import pandas as pd
import xlrd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
import time
import multiprocessing

def Sending_Mail(data):





def main():
  print("-------------------Automation Script : Abhijit Gadhave----------------------")
	print("Script Name : ",argv[0]);
	print("Total Number of Arguments Accepted : ",len(argv))

	if(len(argv)!=2):
		print("Invalid Number of Arguments")
		exit()
	if(argv[1] == '-u' or argv[1] == '-U'):
		print("Usage : It is Used To sending Mail of Happy Birthday Automatically with respective day!!!");
		exit()

	if(argv[1]=='-h' or argv[1]=='-H'):
		print("Help : Name_of_Script Argument")
		print("Argument : Name of that file where data can be stored")
		exit()

	try:
		
		schedule.every(1).minute.do(Receive_Data(argv[1]))

		while(True):
			schedule.run_pending()
			
	except Exception as e:
		print("Exception : ",e);

  


if __name__ == "__main__":
  main()
