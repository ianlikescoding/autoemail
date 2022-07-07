from msilib.schema import RadioButton
import keyboard
import tkinter as tk
import time
import pyperclip as pc
import win32com.client as win32
import os
import pickle
import datetime
# Gmail API utils
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
# for encoding/decoding messages in base64
from base64 import urlsafe_b64decode, urlsafe_b64encode
# for dealing with attachement MIME types
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from mimetypes import guess_type as guess_mime_type

# Request all access (permission to read/send/receive emails, manage the inbox, and more)
SCOPES = ['https://mail.google.com/']
our_email = 'recruiting@phillytech.co'

def gmail_authenticate():
    creds = None
    # the file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # if there are no (valid) credentials availablle, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    #return build('gmail', 'v1', credentials=creds, static_discovery=False)
    return build('gmail', 'v1', credentials=creds)

# get the Gmail API service
service = gmail_authenticate()

window = tk.Tk()
window.title("PhillyTech AutoEmail")
jobList = []
jobNameList = []


f = open("emails.txt", "r")
for line in f:
    if line == "NAME\n":
        jobName = f.readline()
    if line == "SUBJECT\n":
        jobSubject = f.readline()
    if line == "BODY\n":
        jobBody = f.readline() #since the body is almost never 1 line, we have to loop through each line before LINKNAME to add to body 
        while (line != "END\n"):
            line = f.readline()
            if (line == "END\n"): 
                break
            jobBody += line
    if line == "END\n":
        jobArray = [jobSubject.strip(), jobBody.replace("â€™","'")]
        jobList.append(jobArray)
        jobNameList.append(jobName) #this is for the listBox. We put the names in a seprate array so they display cleanly in the ListBox.

temps_var = tk.StringVar(value=jobNameList)
lb = tk.Listbox(
    window,
    listvariable=temps_var,
    height = 5,
    width = 50,
    selectmode='browse'
)

timesList = ["8AM Eastern Time"]
timetoSendLB = tk.Listbox(
    window,

)

greeting = tk.Label(text="Hello, Tkinter")
nameLabel = tk.Label(text="Enter Name")
nameEntry = tk.Entry(width=50)

nameTextEntry = tk.Text(window, width=50, height = 5)

emailLabel = tk.Label(text="Enter Email")
emailEntry = tk.Entry(width=50)

emailTextEntry = tk.Text(window, width=50, height=5)

scheduleSendNum = tk.IntVar(value = 0)
check = tk.Checkbutton(window, text = "Schedule Email? (Currently not working)", variable=scheduleSendNum, onvalue = 1, offvalue = 0)
hourEntry = tk.Entry(width=10)
minuteEntry = tk.Entry(width=10)

AMorPM = tk.StringVar(value = "AM")
AMButton = tk.Radiobutton(window, text="AM", variable = AMorPM, value = "AM")
colonLabel = tk.Label(text=":")
PMButton = tk.Radiobutton(window, text="PM", variable = AMorPM, value = "PM")

GmailOrOutlook = tk.IntVar(value = 0)
GmailButton = tk.Radiobutton(window, text="Send via Gmail", variable = GmailOrOutlook, value = "0")
OutlookButton = tk.Radiobutton(window, text="Send via Outlook", variable = GmailOrOutlook, value = "1")

sameOrNext = tk.StringVar(value = 0)
sameDaySend = tk.Radiobutton(window, text="Send Today", variable = sameOrNext, value = 0)
nextDaySend = tk.Radiobutton(window, text="Send Tomorrow", variable = sameOrNext, value = 1)

send = tk.Button(
    text="Send Email!",
    width = 30,
    height = 1,
)

#functions 

def build_message(destination, obj, body):
    message = MIMEText(body)
    message['to'] = destination
    message['from'] = our_email
    message['subject'] = obj
    return {'raw': urlsafe_b64encode(message.as_bytes()).decode()}


def build_message(destination, obj, body):
    message = MIMEText(body, "html")
    message['to'] = destination
    message['from'] = our_email
    message['subject'] = obj
    return {'raw': urlsafe_b64encode(message.as_bytes()).decode()}

def send_message(service, destination, obj, body):
    return service.users().messages().send(
      userId="me",
      body=build_message(destination, obj, body)
    ).execute()

def sendOutlook(email, subject, body):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = email
    mail.Subject = subject
    mail.HTMLBody = body #this field is optional
    #if (scheduleSendNum.get() == 1):
        #mail.DeferredDeliveryTime = datetime.datetime(2022, )
    mail.Send()

def newSendMail(event):
    if not (lb.curselection()):
        return
    nameText = nameTextEntry.get(0.0,9999.0)
    emailText = emailTextEntry.get(0.0,9999.0)
    emailLine = emailText.split("\n")
    i = 0
    body = str(jobList[int(lb.curselection()[0])][1])
    subject = jobList[int(lb.curselection()[0])][0]
    for line in nameText.split("\n"):
        if line == "":
            return
        name = line
        email = emailLine[i]
        newBody = body.format(name)
        if (GmailOrOutlook.get() == 1):
            sendOutlook(email, subject, newBody)
        else: send_message(service, email, subject, newBody)
        i = i + 1
    nameEntry.delete(0, tk.END)
    emailEntry.delete(0, tk.END)


send.bind("<Button-1>", newSendMail)
lb.pack()
nameLabel.pack()
nameTextEntry.pack()
emailLabel.pack()
emailTextEntry.pack()
check.pack()
hourEntry.pack(side=tk.LEFT)
colonLabel.pack(side=tk.LEFT)
minuteEntry.pack(side=tk.LEFT)
AMButton.pack(side=tk.LEFT)
PMButton.pack(side=tk.LEFT)
sameDaySend.pack(side=tk.LEFT)
nextDaySend.pack(side=tk.LEFT)
GmailButton.pack()
OutlookButton.pack()
send.pack(fill=tk.X, side=tk.BOTTOM)

#From previous versions. Used to automatically set a scheduled time to send a message. No longer implemented.
def scheduleSend(sameOrNextDay, sendHour, sendMinute, AMorPM): #
    keyboard.press_and_release("tab")
    time.sleep(0.10)
    keyboard.press_and_release("tab")
    time.sleep(0.10)
    keyboard.press_and_release("enter")
    time.sleep(0.10)
    keyboard.press_and_release("up")
    time.sleep(0.10)
    keyboard.press_and_release("enter")
    time.sleep(0.10)
    keyboard.press_and_release("up")
    keyboard.press_and_release("enter")
    time.sleep(0.10)
    keyboard.press_and_release("tab")
    time.sleep(0.10)
    if sameOrNextDay == "1":
        keyboard.press_and_release("right") #one more tab than is necessary
        time.sleep(0.10)
    else:
        keyboard.press_and_release("tab")
    time.sleep(0.10)
    keyboard.press_and_release("tab")
    time.sleep(0.10)
    keyboard.press_and_release("tab")
    time.sleep(0.10)
    #breakpoint()
    keyboard.write(sendHour) #write the hour, then the time
    time.sleep(0.10)
    keyboard.write(":")
    time.sleep(0.10)
    keyboard.write(sendMinute)
    time.sleep(0.10)
    keyboard.write(AMorPM)
    time.sleep(0.10)
    keyboard.press_and_release("enter")


tk.mainloop()
#keyboard.press_and_release('alt+tab')


