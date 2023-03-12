import win32com.client
#other libraries to be used in this script
import csv
import os
import re
from datetime import datetime, timedelta
restriction=[]
weight=[]
gross_weight=[]
check_no=0
check_rest=0
messages=[]

print("ABNORMAL LOADS OVERVIEW.")
print("Developed by Kevin Ferguson (kevin.ferguson@durham.gov.uk).")
print("Any HA - Version 1.\n")
inbox = input("Please enter the name of your Abnormal Loads email group folder.  This is how it is displayed in Outlook and is often the email adderss with the @domain.gov.uk removed. E.g. abnormal_loads: ")

#Get the Outlook Session:
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

# open upload and check file for reading.
with open('HWWRoadList.csv', encoding='UTF8') as hww_checkfile:

# read files as csv files
  Check_File = list(csv.reader(hww_checkfile))

#Get messages from Abnomal Loads:
message = mapi.Folders(inbox).Folders("Inbox")
#print(message)
for m in message.Items:
    messages.append(m)
    #messages = mapi.Folders(inbox).Folders("Inbox").Items #Access Abnormal Loads folder

#Restrict to messages from last 24hrs and certain senders:
#received_dt = datetime.now() - timedelta(days=99)
#received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
#messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
#print("5")
#finding all valid emails using regex
#body_content = re.findall(r"[A-Za-z0-9_%+-.]+"
#                 r"@[A-Za-z0-9.-]+"
#                 r"\.[A-Za-z]{2,5}",text)

#Read in messages:
for msg in list(messages):
    body = (msg.body)
    print("\n")
    print (msg.sender)
    #print(msg.senderemailaddress)
    print (msg.ReceivedTime)
    


    # Width and Height are differnet in each ESDAL and ABHAULIER, need to split:
    #sender = msg.senderemailaddress
    #sender_check = re.search("@[\w.]+", str(sender))
    #print(sender_check)

    if "abhaulier.co.uk" in msg.senderemailaddress or "abhaulierservices.co.uk" in msg.senderemailaddress:
        print("ABHAULIER")

        #Regex search for weight - First result is Gross in ESDAL, te is Gross in ABHaulier:
        weight = (re.findall(
            r'((?:[0-9.]{2,7}(?:te| kg)\b))', body))
        try:
            if weight == []:
                print("No weight found")
            else:
                gross_weight = weight[0]
                print ("Weight: " + gross_weight + " (gross)")
                weight=[]
        except:
            print("Error detecting weight")

        #Regex search for Width - " m" in ESDAL (leng, width) "m" in ABHAULIER (leng, proj, proj, width):
        #Width:
        try:
            width = (re.findall(
            r'((?:[0-9.]{2,7}(?:m| m)\b))', body))
            if width == []:
                print("No width found")
            else:
                print ("Width: " + width[1])
                width=[]
        except:
            print("Error detecting width")
            
        #Height:
        try:
            height = (re.findall(
            r'((?:[0-9.]{2,7}(?:m| m)\b))', body))
            if height == []:
                print("No height found")
            else:
                print ("Height: " + height[2])
                height=[]
        except:
            print("Error detecting height")
        
    elif "esdal2.com" in msg.senderemailaddress or "esdal2.co.uk" in msg.senderemailaddress:
        print("ESDAL")

        #Regex search for weight - First result is Gross in ESDAL, te is Gross in ABHaulier:
        weight = (re.findall(
            r'((?:[0-9.]{2,7}(?:te| kg)\b))', body))
        if weight == []:
            print("No weight found")
        else:
            gross_weight = weight[0]
            print ("Weight: " + gross_weight + " (gross)")
            weight=[]
        
        #Regex search for Width - " m" in ESDAL (leng, width) "m" in ABHAULIER (leng, proj, proj, width):
        #ESDAL Uses m for distance on highway also, so need to work in reverse.
        #Width:
        width = (re.findall(
                r'((?:[0-9.]{2,7}(?:m| m)\b))', body))
        if width == []:
            print("No weight found")
        else:
            print ("Width: " + width[-7])
            width=[]
        #Height:
        height = (re.findall(
                r'((?:[0-9.]{2,7}(?:m| m)\b))', body))
        if height == []:
            print("No weight found")
        else:
            print ("Height: " + height[-6])
            height=[]
    else:
        print("Sender not Esdal or Abhaulier.")
    
    #Check to see if any streets matching restrictions are in the message body:
    for i in Check_File:
        check_length= len(Check_File)
        if i[0] in body:
            #restriction.append(msg.ReceivedTime)
            #restriction.append(msg.sender)
            restriction.append(i)
            print (restriction)
            restriction = []
        check_no += 1
        
        if check_no == check_length:
            print ("END OF RESTRICTIONS CHECK.")
            check_no=0
            
input()
