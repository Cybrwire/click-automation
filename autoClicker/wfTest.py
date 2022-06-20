import os
import pyautogui as pag
from pynput import mouse, keyboard

# variables
screenSize = pag.size()
images = []
kb = keyboard.Controller()
keyPresses = ['Key.enter','Key.end','Key.tab']
imgDirectory = '../images'

# fill array with filenames from images directory
for file in os.listdir(imgDirectory): 
    images.append(file)

#center point of search image
a,b = pag.locateCenterOnScreen(images[0])


print('x: {0}, y: {1}'.format(a,b))
if pag.onScreen(a,b):
    pag.moveTo(a,b,1)
    pag.click()
    print('pass')
else:
    print('fail')
    print(screenSize)