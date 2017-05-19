import numpy as np
import cv2
import time
from skimage.measure import structural_similarity as ssim
import matplotlib.pyplot as plt
import pytesseract
from pytesser import *
from PIL import Image, ImageEnhance, ImageFilter
import winsound
import sys
import imaplib
import getpass
import email as em
import datetime
import os
import ImageGrab
import smtplib
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pyttsx
import win32evtlog
import win32api
import win32con
import win32security # To translate NT Sids to account names.
import win32evtlogutil
from twilio.rest import Client

def ReadLog(computer, logType="Application", dumpEachRecord = 0):
    # read the entire log back.
    h=win32evtlog.OpenEventLog(computer, logType)
    numRecords = win32evtlog.GetNumberOfEventLogRecords(h)
#       print "There are %d records" % numRecords
    logToText=""
    num=0
    while 1:
        objects = win32evtlog.ReadEventLog(h, win32evtlog.EVENTLOG_BACKWARDS_READ|win32evtlog.EVENTLOG_SEQUENTIAL_READ, 0)
        if not objects:
            break
        for object in objects:
            # get it for testing purposes, but dont print it.
            msg = win32evtlogutil.SafeFormatMessage(object, logType)
            if object.Sid is not None:
                try:
                    domain, user, typ = win32security.LookupAccountSid(computer, object.Sid)
                    sidDesc = "%s/%s" % (domain, user)
                except win32security.error:
                    sidDesc = str(object.Sid)
                user_desc = "Event associated with user %s" % (sidDesc,)
            else:
                user_desc = None
            if dumpEachRecord:
                logToText+="\n"+"Event record from %r generated at %s" % (object.SourceName, object.TimeGenerated.Format())
                if user_desc:
                    logToText+="\n"+user_desc
                try:
                    logToText+="\n"+msg
                except UnicodeError:
                    logToText+="\n"+"(unicode error printing message: repr() follows...)"
                    logToText+="\n"+repr(msg)
        num = num + len(objects)
    if numRecords == num:
        logToText+="\n"+"Successfully read all "+ str(numRecords)+ " records"
    else:
        logToText+="\n"+"Couldn't get all records - reported %d, but found %d" % (numRecords, num)
        logToText+="\n"+"(Note that some other app may have written records while we were running!)"
    win32evtlog.CloseEventLog(h)
    return(logToText)

# this function take a message and convert to speech using pyttsx
def sayCommand(message):
	eng=pyttsx.init()
	eng.say(message)
	eng.runAndWait()

def sendEmail(sendTo,textfile,logfile,img):
    """Retrieves the error.txt and an the taken image and sends an email
    with those attached"""
    # Open a plain text file for reading
    msg = MIMEMultipart()

    # Read the text file <-- Error msg from OCR module
    if(textfile!=""):
        fp = open(textfile, 'rb')
        text = MIMEText(fp.read())
        fp.close()
        msg.attach(text)

    if(logfile=='y'):
        filename = "log.txt"
        fp = open(filename)
        log = MIMEText(fp.read())
        fp.close()
        log.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(log)

    msg['Subject'] = 'An event has occurred at the MS'
    msg['From'] = "mass.checker@gmail.com"
    msg['To'] = sendTo

    # Load screenshot and attach to email
    fp = open(img, 'rb')
    img = MIMEImage(fp.read())
    fp.close()
    msg.attach(img)

    # Send the message
    server = smtplib.SMTP('smtp.gmail.com',587)
    server.starttls()
    server.login("mass.checker@gmail.com", "massspecchecker1234")

    server.sendmail("mass.checker@gmail.com", sendTo, msg.as_string())
    server.quit()

def getCommand():
    """Retrieves the command specified by the user through email"""
    command="error"
    M = imaplib.IMAP4_SSL('imap.gmail.com')
    try:
        M.login('mass.checker@gmail.com', 'massspecchecker1234')
    except imaplib.IMAP4.error:
        print "LOGIN FAILED!!! "
    rv, mailboxes = M.list()
    rv, data = M.select("INBOX")
    if rv == 'OK':
        print "Waiting for command ...\n"
        rv, data = M.search(None, "(UNSEEN)")
    if rv != 'OK':
        print "No messages found!"
        return
    for num in data[0].split():
        rv, data = M.fetch(num, '(RFC822)')
        if rv != 'OK':
            print "ERROR getting message", num
            return
        msg = em.message_from_string(data[0][1])
        command=msg['Subject']
        M.close()
        M.logout()
        if(command is None):
            command="error"
        return(command)

def getInput(inputType):
    inputed = raw_input("%s:" % inputType)
    return inputed

def mse(imageA, imageB):
    """Calculates the Mean Square Error between two images"""
    # the 'Mean Squared Error' between the two images is the
    # sum of the squared difference between the two images;
    # NOTE: the two images must have the same dimension
    err = np.sum((imageA.astype("float") - imageB.astype("float")) ** 2)
    err /= float(imageA.shape[0] * imageA.shape[1])

    # return the MSE, the lower the error, the more "similar"
    # the two images are
    return err

def compare_images(imageA, imageB):
    """Compares two images"""
    # compute the mean squared error and structural similarity
    # index for the images
    m = mse(imageA, imageB)
    s = ssim(imageA, imageB)
    return s

def getScreen():
    """Takes a screenshot and saves it as screenshot.png in the current
    working directory"""
    img=ImageGrab.grab()
    img.save("screenshot.png")
    return("screenshot.png")

def webcameCapture():
    """Takes a photo through the webcam and saves it as webcam.png in
    the current working directory"""
    retval, frame = cap.read()
    cv2.imwrite("webcam.png",frame)
    img=cv2.imread("webcam.png")
    return(img)

def enhanceImage(img):
    """Enhances an image"""
    mg = Image.open(img)
    mg.load()
    #mg = mg.filter(ImageFilter.MedianFilter())
    #enhancer = ImageEnhance.Contrast(mg)
    #mg = enhancer.enhance(2)
    #mg = mg.convert('1')
    #mg.save('out.jpg')
    return(mg)

email="payam.emami@medsci.uu.se;stephanie.herman@medsci.uu.se"
interval=2
alarm="off"
webcamOrScreen="s"
OCROoption="on"
threshold=0.7
logfile="n"
commandEmail="y"
mobile=""

#### reading required inputs from user
if(len(sys.argv)==1):
    email=getInput("Please enter your email address")
    mobile=getInput("Please provide mobile number if you want alerts through sms")
    interval=int(getInput("Please enter the time interval for taking photos"))
    alarm=getInput("Please enter [on/off] for turning alarm system on or off")
    webcamOrScreen=getInput("Please enter [w/s] for reading from webcam or screen")
    OCROoption=getInput("Please enter [on/off] if you want to turn OCR on or off")
    threshold=float(getInput("Please enter the similarity threshold for alerting. It should be between 0 and 1 where 0 is least and 1 is most sensitive"))
    logfile=getInput("Do you want to include the MS log file? [y/n]")
    commandEmail=getInput("Do you want to be able to control the MS computer through email? [y/n]")
if(len(sys.argv)>1):
    print("Using default settings! The email will be sent to %s"% email)
#### in the begining we don't have any image
preImage=None
#### this handle for webcam!
cap=""
#### check if user wants webcam ? if yes, turn it on.
if(webcamOrScreen=="w"):
    print("Turning on the webcam")
    cap = cv2.VideoCapture(0)

while(True):
    # wait for the interval user selected
    time.sleep(interval)
    # if user selected screenshot, take one!
    if(webcamOrScreen=="s"):
        print("Getting screenshot ...")
        image=img=cv2.imread(getScreen())
    else:# otherwise, we use webcam
        print("Taking photo from webcam ...")
        image=webcameCapture()
    print("Transforming the image to grayscale ...")
    image=cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    # write the image as tmp
    cv2.imwrite("grayscale.png",image)
    # when the program starts we don't have any image to compare with. so just set the previous image to the new one
    if(preImage is None):
        preImage=image
    else:
        newImage=image
        error=compare_images(preImage, newImage)
        print(error)
        if(error<threshold):
            logTxt=""
            errorText=""
            print("The screen has changed! preparing the report ...")
            if(OCROoption=="on"):
                print("Trying to detect possible error ...")
                pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract'
                print("text: ")
                errorText=pytesseract.image_to_string(enhanceImage("grayscale.png"))
                print(errorText)
            if(errorText==""):
                errorText="An event has occurred at the MS. Please check the attachment"
            if(logfile=="y"):
               print("Reading log ...")
               logTxt=ReadLog(None,"Application",1)
               file = open("log.txt","w")
               file.write(logTxt.encode('utf8'))
               file.close()
            print("Emailing ....")
            file = open("error.txt","w")
            file.write(errorText)
            file.close()
            sendEmail(email,"error.txt",logfile,"grayscale.png")
            if(mobile!=""):
                # Find these values at https://twilio.com/user/account
                account_sid = "AC04bfb7de2206875f4427f2eea454e49a"
                auth_token = "86f75bf437606bac2924e54ef3f76a7f"
                client = Client(account_sid, auth_token)

                message = client.api.account.messages.create(to=mobile,
                                             from_="+46769437491",
                                             body="Check you email! The MS is experiencing problems.")
            if(alarm=="on"):
                winsound.PlaySound('alarm.wav', winsound.SND_FILENAME | winsound.SND_ASYNC | winsound.SND_LOOP)
            if(commandEmail=="y"):
                while(True):
                     time.sleep(5)
                     command=getCommand()
                     if(command is None):
                         command=""
                     if(command=="exit"):
                         print("Executing command ...")
                         sys.exit()
                     if(command=="continue"):
                         print("Executing command ...")
                         break;
                     if(command=="reset"):
                         print("Executing command ...")
                         preImage="None"
                         break;
                     if(command=="shutdown"):
                         print("Executing command ...")
                         os.system('shutdown -s')
                     if(command=="alarm"):
                         print("Executing command ...")
                         winsound.PlaySound('alarm.wav', winsound.SND_FILENAME | winsound.SND_ASYNC | winsound.SND_LOOP)
                     if(command=="stopalarm"):
                         print("Executing command ...")
                         winsound.PlaySound(None, winsound.SND_ASYNC)
                     if(command=="status"):
                         print("Executing command ...")
                         if(webcamOrScreen=="s"):
                             getScreen()
                             sendEmail(email,"",logfile,"screenshot.png")
                         if(webcamOrScreen=="w"):
                             webcameCapture()
                             sendEmail(email,"",logfile,"webcam.png")
                     if "say" in command:
                         # the format of command for this section is "say:message:x". The program says the message x times.
                         # split the message by ":"
                         commandSplit=command.split(":")
                         # second element is message
                         message=commandSplit[1]
                         # and third element is repetitions
                         repetitions=int(commandSplit[2])
                         print("Executing command ...")
                         # then we say the message
                         for x in range(0, repetitions):
                             sayCommand(message)
                     if(command=="log"):
                             print("Reading log ...")
                             logTxt=ReadLog(None,"Application",1)
                             print("writing log to text file ...")
                             file = open("log.txt","w")
                             file.write(logTxt.encode('utf8'))
                             file.close()
                             print("Executing command ...")
                             sendEmail(email,"error.txt","y","grayscale.png")

            preImage=newImage
if(webcamOrScreen=="w"):
    cap.release()
    cv2.destroyAllWindows()
