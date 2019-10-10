import win32com.client as win32
import psutil
import os
import subprocess
import openpyxl
from pathlib import Path
from datetime import datetime, date
#import win32api

today = date.today()
now = datetime.now()

#Create and Set the year, month, and day variables
year = now.strftime("%Y")
month = now.strftime("%m")
day = now.strftime("%d")
monday = int(day) - 4 #created the variable for monday (4 days before)
mondayString = str(monday)
mondayFullDate = month + '/' + mondayString + '/' + year
weekOf = month + '_' + mondayString + '_' + year
extend = month + '_' + mondayString + '_' + year + ".xlsx"
filename = "C:\\Users\\YourDesiredNewFileName" + extend

#Opens the XLSX Workbook and assigns appropriate cell values
def fileUpdate():
    print("Loading Workbook")
    wb = openpyxl.load_workbook('C:\\Users\\YourPathHere') #Path to reference file for update
    sheet = wb['Sheet1'] #Selects the sheet
    print("Updating Date")
    sheet['B9'] = mondayFullDate #Sets the date to Monday (4 days earlier) on cell B9
    dayOff = input("Did you have a day off? If so, what type of day (Enter 0, S, J, BH, or V) or 'n' for none: ")
    dayOff = dayOff.lower()
    answers = ['0', 's', 'j', 'bh', 'v']
    none = ['n', 'none']
    if dayOff not in answers:
        print("Please enter a valid response")
    elif dayOff in none:
        print("here")
        #print(sheet['B9'].value)
        #print(fileName)
        print("Saving new file...")
        wb.save(filename)
        print("Save Succesful!")
        sendNotification()
    elif dayOff in answers:
        print("in dayoff .lower")
        the_day_off = input("Which day? M, T, W, Th, F: ")
        if the_day_off.lower() == 'm':
            sheet['C15'] = dayOff.upper()
            print("Saving new file...")
            wb.save(filename)
            print("Save Succesful!")
            sendNotification()
        elif the_day_off.lower() == 't':
            sheet['C16'] = dayOff.upper()
            print("Saving new file...")
            wb.save(filename)
            print("Save Succesful!")
            sendNotification()
        elif the_day_off.lower() == 'w':
            sheet['C17'] = dayOff.upper()
            print("Saving new file...")
            wb.save(filename)
            print("Save Succesful!")
            sendNotification()
        elif the_day_off.lower() == 'th':
            sheet['C18'] = dayOff.upper()
            print("Saving new file...")
            wb.save(filename)
            print("Save Succesful!")
            sendNotification()
        elif the_day_off.lower() == 'f':
            sheet['C19'] = dayOff.upper()
            print("Saving new file...")
            wb.save(filename)
            print("Save Succesful!")
            sendNotification()

# Drafting and sending email notification to senders.
# You can add other senders' email in the list
def sendNotification():
    const = win32.constants
    olMailItem = 0x0
    obj = win32.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = "ENTER MAIL SUBJECT HERE" + month + '/' + mondayString + '/' + year
    newMail.Body = ''
    newMail.BodyFormat = 2 # olFormatHTML https://msdn.microsoft.com/en-us/library/office/aa219371(v=office.11).aspx
    newMail.HTMLBody = "<HTML><BODY>Hi,<br><br>Enter Your body text here <span style='font-weight: bold; text-decoration: underline'>Bold and Underline</span> More body text.<br><br>Thanks,<br>Your Name</BODY></HTML>" #"<HTML><BODY>Enter \nthe <span style='color:red'>message</span> text here.</BODY></HTML>"
    newMail.To = "to@example.com"
    newMail.CC = "cc@example.com"
    attachment1 = filename
    print("Attaching", attachment1)
    newMail.Attachments.Add(Source=attachment1)
    print("Attachment Added")
    print("Opening message...")
    newMail.display(True)
    #newMail.send()
    print("Mail Successfully Sent!")
    quit()

# Open Outlook.exe. Path may vary according to system config
# Please check the path to .exe file and update below
  
'''def openOutlook():
    print("Checking if Microsoft Outlook is open...")
    try:
        subprocess.call(['C:\Program Files\Microsoft Office\Office15\Outlook.exe'])
        os.system("C:\Program Files\Microsoft Office\Office15\Outlook.exe");
    except:
        print("Outlook didn't open successfully")
        print("Trying again...")
 
# Checking if outlook is already opened. If not, open Outlook.exe and send email
for item in psutil.pids():
    p = psutil.Process(item)
    if p.name() == "OUTLOOK.EXE":
        flag = 1
        break
    else:
        flag = 0
 
if (flag == 1):
    fileUpdate() #was sendNotification()
else:
    openOutlook()
    sendNotification()

fileUpdate()'''

def main():
    #openOutlook()
    sendNotification()

if __name__ == "__main__":
    main()
