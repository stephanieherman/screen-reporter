# screen-reporter
In this page we provide a screen monitoring application which was developed as a course project in Advanced Scientific Programming in Python - Uppsala University.

The application is build to monitor a screen (through screenshots) or an environment (through the webcam). The application will take screenshots/photos in a certain interval and then calculate the similarity between the two following images. If the similarity is less than a user specified threshold, an email will be sent to the user with the latest screenshot/photo attached to the email and, if specified, also the log file and the error message. An alarm system can also be used to alert the surroundings and a mobile number can be given where a notification telling you to check your email will be sent.

A feature has also been added providing the user with the possibility to control the host computer through email, by sending code words to `mass.checker@gmail.com`. You also have a possibility to send messages that will be shown in the command line running the code.

## Run the application
To be able to run this application, several python packages has be in stalled on your in-house python environment. You should be able to import the following:

```python
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
```

After successful installation of all dependencies, you can run the script in your command line as following:

```bash
python screenreporter.py
```

**NOTE**

You have to be in the same directory as you cloned the git library

## Parameters
You will then be asked to set a couple of parameters:

```
Please enter your email address:
Please provide mobile number if you want alerts through sms:
Please enter the time interval for taking photos:
Please enter [on/off] for turning alarm system on or off:
Please enter [w/s] for reading from webcam or screen:
Please enter [on/off] if you want to turn OCR on or off:
Please enter the similarity threshold for alerting. It should be between 0 and 1 where 0 is least and 1 is most sensitive:
Do you want to include the MS log file? [y/n]:
Do you want to be able to control the MS computer through email? [y/n]:
```
