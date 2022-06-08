# Author: Justin Montgomery
# Date: 05/11/2022
# Version: Python 3.8.9
# Purpose: This module allows the user to select two points on the screen creating a box around an area on screen to be searched for.

import pyautogui as pag
from pynput import mouse
import time

screenSize = [0,0,1919,1079]
currentItem = []#[x1,y1,x2,y2] 
searchItemsList = []
looping = True #controls while loop in main()

def onClick(x,y,button,pressed):
    
    #end item selection
    if button == mouse.Button.right:
        listener.stop()
        
    #selected area must be within screen size
    elif (0 < x < 1919) and (0 < y < 1079) and pressed and button==mouse.Button.left:
        
        if len(currentItem) <= 2:
            currentItem.append(x)
            currentItem.append(y)
            print('collected first point')
        #stop listener when we get two sets of coordinates
        if len(currentItem) == 4:
            listener.stop()
            target = selectItem(currentItem)
            searchItemsList.append(target)
            print('collected second point')
            currentItem.clear()
            setMousePosition(target)

#returns center coordinates of an area selected by user 
def selectItem(currentItem): 

    #area to be screenshotted(left, top, width, height)
    searchObject = pag.screenshot(region=(currentItem[0],currentItem[1],currentItem[2]-currentItem[0],currentItem[3]-currentItem[1]))
    targetItem = pag.locateCenterOnScreen(searchObject, grayscale=False,confidence=0.9)
    return targetItem

def setMousePosition(target):
    x, y = target[0], target[1]
    print('x= {}, y= {}', x,y)
    pag.moveTo(x,y)
    
print('start')
with mouse.Listener(on_click=onClick) as listener:
    listener.join()
    print('listener started')

