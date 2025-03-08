#all import functions
import time
import cv2
import os
import pytesseract
import numpy as np
import pyautogui
import win32gui
from PIL import Image
import math
import win32.lib.win32con as win32con
import nltk

#tells code where to find Tessaract program, if your drive isn't C this needs to be changed.
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract"
TESSDATA_PREFIX = 'C:/Program Files (x86)/Tesseract-OCR'

#Gains information for the program to execute
pagename = input("Please type in the name of the application that the screenshare is going through.\n")
if pagename =="":
    pagename = "SM-G990U1"
hwnd = win32gui.FindWindow(None,pagename)
calibrationval = input("Type Y and hit enter to calibrate manually. Type anything else to test auto calibration (not reccomended).\n")

#inital variables needed for calibration to work
calibrated = "s"
size = pyautogui.size()
resx, resy = size
r,g,b = 255, 184, 26


#calibration loop
if calibrationval == "y":
    while calibrated != "y":
        try:
            print("You should download a software that tells you your mouse position in coordinates. For an example, look up Mpos.\nNow open a screenshot of the ranked draft screen, and full screen your screenshare window for best results.")
            x = int(input("Visualize a tight box around the map title. Enter the x coordinate of the top left corner of the box.\n"))
            y = int(input("Now enter the y value of the same point\n"))
            bx = int(input("Do the same with the bottom right of the box, making sure to just enter the x coordinate\n"))
            by = int(input("Now enter the y value of the same point\n"))
            width = abs(bx-x)+int(resx*50/1920)
            height = abs(by-y)
            checkx = int(input("Now locate the golden progress bar at the top of the page. \nPlace your mouse in the first circle, making sure it is pointing in the center of the circle.\nType in the x coordinate of this point.\n"))
            checky = int(input("Now report the y coordinate of this point\n"))

            while calibrated == "s":
                valuenotneeded = input("Make sure your phone is on and the screen is shared to your PC. \nThen hit enter, wait for an image to appear, then open this command tab.\n")
                win32gui.SetForegroundWindow(hwnd)
                checkshot = pyautogui.screenshot()
                checker = checkshot.getpixel((checkx, checky))
                r,g,b = checker[0],checker[1],checker[2]
                screenshot = pyautogui.screenshot(region=(x, y, width, height))
                screenshot.show()
                time.sleep(5)
                calibrated = input("Type in Y if you are able to read the entire map name and there is empty space at the end of the box. \nTo see the screenshot again type S. \nIf you are unable to read the screenshot, hit enter.\n")
                calibrated = calibrated.lower()
        except:
            print("Something went wrong, please try again!\n")
            time.sleep(5)
            continue

#auto calibration (calibrated to 1920*1080 and using scrcpy on full screen with Galaxy S21 FE)
else:
    x,y = int(0.09375*resx),int(0.157407407*resy)
    width,height = int(0.114583333*resx), int(0.0277777778*resy)
    checkx, checky= int(0.403125*resx), int(0.114814815*resy)

#more variables needed for program
extra = ["[","]","|"," ","{","}","(",")","'",";",":",".",",","?","!","\\","/","\""]
exlen = len(extra)

#searches the program directory for a folder named "SpenLC". Feel free to rename your folder and change this code.
#Make sure that this folder and images are present. Code won't work without them.
all_entries = os.listdir("SpenLC")
file_names = [entry for entry in all_entries if os.path.isfile(os.path.join("SpenLC", entry))]


#If I missed a file type feel free to add it.
file_type =[".png", ".jpeg", ".webp", ".avif", ".gif"]

#adjustments to image for readability
def correction1(img):
    grayimage = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    product = cv2.threshold(grayimage, 0, 255, cv2.THRESH_TOZERO + cv2.THRESH_OTSU)[1]
    return product

def correction2(img):
    kernel = np.ones((2, 2), np.uint8)
    erodeimage = cv2.erode(img, kernel, iterations=1)
    grayimage = cv2.cvtColor(erodeimage, cv2.COLOR_BGR2GRAY)
    product = cv2.threshold(grayimage, 0, 255, cv2.THRESH_TOZERO + cv2.THRESH_OTSU)[1]
    return product

#Uses Pytessaract to read the image
def reading(img):
    proyour = (pytesseract.image_to_string(img, lang='eng',
                                           config='--psm 7 -c tessedit_char_whitelist= QWERTYUIOPASDFGHJKLZXCVBNMqwertyuiopasdfghjklzxcvbnm'))
    for i in extra:
        proyour = proyour.replace(i, "")
    proyour = proyour.lower()
    proyour = proyour.strip()
    # apparently "Open Business" reads as @hanesinass and I have no idea how to fix it
    if proyour == "@hanesinass":
        proyour = "openbusiness"
    print(proyour)
    return proyour

#prints image file if map name detected
def print_image(mapname):
    im = cv2.imread("SpenLC/" + mapname)
    im = cv2.resize(im, (960, 540))
    cv2.imshow("MapGuide", im)
    mapper = win32gui.FindWindow(None, "MapGuide")
    win32gui.SetForegroundWindow(mapper)
    while True:
        cv2.imshow("MapGuide", im)
        cv2.waitKey(1000)
        checkshot = pyautogui.screenshot()
        checker = checkshot.getpixel((checkx, checky))
        if checker[0]!=255:
            cv2.destroyAllWindows()
            break

#Start of loop briefing
print("The program will constantly print the rgb values of the location where the progress bar is. \nThe program should attempt to read the title of the map if the color values return the same as the progress bar. \nIf this does not work, you may have to recalibrate the program (restart it)")
time.sleep(10)

#start while loop (Will run until program terminated with "Ctrl+C")
while True:
   try:
    #set Phone screen share to top of screen.
       win32gui.SetForegroundWindow(hwnd)
   except:
       print("Could not find window")
    #Checks if draft screen is live
   checkshot = pyautogui.screenshot()
   checker = checkshot.getpixel((checkx, checky))
   print(checker)
   if checker[0] == r and g-5<checker[1]<g+5 and b-5 <checker[2]< b+5:

       #takes picture of title and refractors it for better reading
       screenshot = pyautogui.screenshot(region=(x, y, width, height))
       new_size = (screenshot.width * 2, screenshot.height * 2)
       screenshot = screenshot.resize(new_size, Image.LANCZOS)
       screenshot.save("name.png")

       #performs three different alterations for best chance of reading title
       image = cv2.imread('name.png')
       name1 = reading(image)
       product = correction1(image)
       name2 = reading(product)
       product2 = correction2(image)
       name3 = reading(product2)

       #attempts to match title to a file name
       namelist = [name1,name2,name3]
       found = False
       for names in namelist:
           if found == True:
               break
           try:
               #protocol that adjusts for slight misreads by tessaract OCR
               print("Engaging Levenshtein protocol")
               for z in file_names:
                   for d in file_type:
                       c = z.replace(d, "")
                   similar = nltk.edit_distance(names, c)
                   if similar < 5:
                       print_image(z)
                       found = True
                       break
           except:
               print("unable to read screenshot")
   time.sleep(1)
