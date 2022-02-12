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
	
	with open("Birth_Image.jpg","rb") as f:
		img_data = f.read();

	msg = MIMEMultipart()
	msg["From"] = "abhigadhave97@gmail.com"
	msg["To"] = data[4]
	msg['subject'] = "Happy Birthday(Wishes)"
	body = "Hello "+data[1]+",\n"+"Wishing You a Very Very Wonderful birthday! May all your dreams comes true!!Have a great year ahead.💫️🤗️💥️🍫️💥️💥️"
	msg.attach(MIMEText(body, 'plain'))

	image = MIMEImage(img_data,name="Birth_Image")
	msg.attach(image)

	s = smtplib.SMTP('smtp.gmail.com', 587)
	s.ehlo()
	s.starttls()
	s.ehlo()
	
	s.login("abhigadhave97@gmail.com", "Abhijitkg111")
	
	text = msg.as_string()
	s.sendmail("abhigadhave97@gmail.com", data[4], text)
	
	s.quit()

def Receive_Data(strX):
	todayDate = date.today()
	
	stringD = (str)(todayDate)
	x = stringD.split("-")
	
	DOB = x[1]+"-"+x[2]
	
	workbook = pd.read_excel(strX)
	
	for i in range(0,len(workbook.index)):
		if(DOB == workbook['DOB'].iloc[i])
		
		wb = xlrd.open_workbook(strX)
		sheet = wb.sheet_by_index(0)
		data.append(sheet.row_values(i+1))
		del wb
		
	pool = multiprocessing.pool()
	pool.map(Sending_Mail,data)
	pool.close()





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
