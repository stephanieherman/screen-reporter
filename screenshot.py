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

def getCommand():
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
	# the 'Mean Squared Error' between the two images is the
	# sum of the squared difference between the two images;
	# NOTE: the two images must have the same dimension
	err = np.sum((imageA.astype("float") - imageB.astype("float")) ** 2)
	err /= float(imageA.shape[0] * imageA.shape[1])
	
	# return the MSE, the lower the error, the more "similar"
	# the two images are
	return err
 
def compare_images(imageA, imageB, title):
	# compute the mean squared error and structural similarity
	# index for the images
	m = mse(imageA, imageB)
	s = ssim(imageA, imageB)
	return s


#import pyscreenshot as ImageGrab
import ImageGrab
def getScreen():
    img = ImageGrab.grab()
    img.save('screenshot2.jpg')
    img=cv2.imread("screenshot2.jpg")
    return img
def webcameCapture():
    retval, frame = cap.read()
    cv2.imwrite("screenshot2.jpg",frame)
    img=cv2.imread("screenshot2.jpg")
    return(img)
def enhanceImage(img):
    mg = Image.open(img)
    mg.load()
    #mg = mg.filter(ImageFilter.MedianFilter())
    #enhancer = ImageEnhance.Contrast(mg)
    #mg = enhancer.enhance(2)
    #mg = mg.convert('1')
    #mg.save('out.jpg')
    return(mg)


#### reading required inputs from user
email=getInput("Please enter your email address ")
interval=int(getInput("Please enter the time interval of taking photo"))
alarm=getInput("Please enter on/off for turning alarm system on or off ")
webcamOrScreen=getInput("Please enter w/s for reading from webcam or screen ")
OCROoption=getInput("Please enter on/off if you want to turn OCR on or off")
threshold=float(getInput("Please enter the threshold for alarm. It should be between 0 and 1 where 0 is least and 1 is most sensitive"))
commandEmail=getInput("Drive using email ? y/n ")
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
        image=getScreen()
    else:# otherwise, we use webcam
        print("Taking photo from webcam ...")
        image=webcameCapture()
    print("Transforming the image to grayscale ...")
    image=cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    # write the image as tmp
    cv2.imwrite("tmp.png",image)
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
                errorText=pytesseract.image_to_string(enhanceImage("tmp.png"))
                print(errorText)
            if(errorText==""):
                errorText="There is something happening with MS. Please check the attachment"
            print("Emailing ....")
            file = open("error.txt","w")
            file.write(errorText)
            file.close() 
            if(alarm=="on"):
                winsound.Beep(300,2000)
                winsound.Beep(600,2000)
                winsound.Beep(800,2000)
            if(commandEmail=="y"):
                while(True):
                     time.sleep(5)
                     command=getCommand()
                     if(command=="exit"):
                         exit()
                     if(command=="continue"):
                         break;
                     if(command=="reset"):
                         preImage="None"
                         break;
                     if(command=="shutdown"):
                         os.system('shutdown -s')
            preImage=newImage
if(webcamOrScreen=="w"):
    cap.release()
    cv2.destroyAllWindows()
