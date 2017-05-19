import numpy as np
import cv2
import time
from skimage.measure import structural_similarity as ssim
import matplotlib.pyplot as plt
#import pytesseract
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

def compare_images(imageA, imageB, title):
    """Compares two images"""
    # compute the mean squared error and structural similarity
    # index for the images
    m = mse(imageA, imageB)
    s = ssim(imageA, imageB)
    return s

#import pyscreenshot as ImageGrab
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


#### reading required inputs from user
email=getInput("Please enter your email address")
interval=int(getInput("Please enter the time interval for taking photos"))
alarm=getInput("Please enter [on/off] for turning alarm system on or off")
webcamOrScreen=getInput("Please enter [w/s] for reading from webcam or screen")
OCROoption=getInput("Please enter [on/off] if you want to turn OCR on or off")
threshold=float(getInput("Please enter the similarity threshold for alerting. It should be between 0 and 1 where 0 is least and 1 is most sensitive"))
logfile=getInput("Do you want to include the MS log file? [y/n]")
commandEmail=getInput("Do you want to be able to control the MS computer through email? [y/n]")
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
        print(1)
    else:
        newImage=image
        error=compare_images(preImage, newImage, "Original vs. Original")
        print(error)
        if(error<threshold):
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
            print("Emailing ....")
            file = open("error.txt","w")
            file.write(errorText)
            file.close()
            sendEmail(email,"error.txt",logfile,"grayscale.png")
            if(alarm=="on"):
                winsound.PlaySound('alarm.wav', winsound.SND_FILENAME | winsound.SND_ASYNC | winsound.SND_LOOP)
            if(commandEmail=="y"):
                while(True):
                     time.sleep(5)
                     command=getCommand()
                     if(command=="exit"):
                         sys.exit()
                     if(command=="continue"):
                         break;
                     if(command=="reset"):
                         preImage="None"
                         break;
                     if(command=="shutdown"):
                         os.system('shutdown -s')
                     if(command=="alarm"):
                         winsound.PlaySound('alarm.wav', winsound.SND_FILENAME | winsound.SND_ASYNC | winsound.SND_LOOP)
                     if(command=="stopalarm"):
                         winsound.PlaySound(None, winsound.SND_ASYNC)
                     if(command=="status"):
                         if(webcamOrScreen=="s"):
                             getScreen()
                             sendEmail(email,"",logfile,"screenshot.png")
                         if(webcamOrScreen=="w"):
                             webcameCapture()
                             sendEmail(email,"",logfile,"webcam.png")

            preImage=newImage
if(webcamOrScreen=="w"):
    cap.release()
    cv2.destroyAllWindows()
