# Author: Justin Montgomery
# Date: 05/11/2022
# Version: Python 3.10.5
# Purpose: This module allows the user to select two points on the screen creating a box around an area on screen to be searched for.

import pyautogui as pag
from pynput import mouse, keyboard

# variables
screenSize = pag.size()
currentItem = []#[x1,y1,x2,y2] 
searchItemsList = []
kb = keyboard.Controller()


def onClick(x,y,button,pressed):
    
    #end item selection
    if button == mouse.Button.right:
        listener.stop()
        
    #selected area must be within screen size
    elif pag.onScreen(x,y) and pressed and button==mouse.Button.left:
        
        if len(currentItem) <= 2:
            currentItem.append(x)
            currentItem.append(y)
            if len(currentItem) == 2: 
                print('collected first point')
        #stop listener when we get two sets of coordinates
        if len(currentItem) == 4:
            listener.stop()
            target = selectItem(currentItem)
            searchItemsList.append(target)
            print('collected second point')
            currentItem.clear()
            setMousePosition(target)

#returns center coordinates between two points selected by user 
def selectItem(currentItem): 

    #area to be screenshotted(left, top, width, height)
    searchObject = pag.screenshot(region=(currentItem[0],currentItem[1],currentItem[2]-currentItem[0],currentItem[3]-currentItem[1]))
    targetItem = pag.locateCenterOnScreen(searchObject, grayscale=False,confidence=0.9)
    return targetItem

def setMousePosition(target):
    x, y = target[0], target[1]
    print('x= {0}, y= {1}'.format(x,y))
    pag.moveTo(x,y)
    


print('start')

with mouse.Listener(on_click=onClick) as listener:
    listener.join()
