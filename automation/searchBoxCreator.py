#!/usr/bin/env python
# Author: Justin Montgomery
# Date: 05/11/2022
# Version: Python 3.10.5
# Purpose: This module allows the user to select two points on the screen
# creating a box around an area on screen to be searched for.

import pyautogui as pag
from pynput import mouse, keyboard
import sys

def python_310():
    """Returns whether the current python version is above or equal to 3.10"""
    return sys.version_info >= (3, 10)

def selectItem(currentItem):
    """Returns the center coordinates between two points selected by user"""
    #area to be screenshotted(left, top, width, height)
    searchObject = pag.screenshot(region=(currentItem[0],currentItem[1],
            currentItem[2]-currentItem[0],currentItem[3]-currentItem[1]))
    return pag.locateCenterOnScreen(searchObject, grayscale=False,confidence=0.9)

def setMousePosition(target):
    """Sets the mouse position to the center of the target"""
    x, y = target[0], target[1]
    print(f'x= {x}, y= {y}')
    pag.moveTo(x,y)

def onClick(x,y,button,pressed):
    """Called when a mouse button is pressed"""
    #end item selection
    if button == mouse.Button.right:
        listener.stop()

    #selected area must be within screen size
    elif pag.onScreen(x,y) and pressed and button==mouse.Button.left:
        l = len(currentItem)
        if python_310():
            match [l<=2, l==4]:
                case (True, False):
                    currentItem.append((x,y))
                    if currentItem == 2:
                        print('collected first point')
                case (False, True):
                    listener.stop()
                    target = selectItem(currentItem)
                    searchItemsList.append(target)
                    print('collected second point')
                    currentItem.clear()
                    setMousePosition(target)
        else:
            if l<=2:
                currentItem.append((x,y))
                if currentItem == 2:
                    print('collected first point')
            elif l==4:
                listener.stop()
                target = selectItem(currentItem)
                searchItemsList.append(target)
                print('collected second point')
                currentItem.clear()
                setMousePosition(target)

if __name__ == '__main__':
    # declare variables
    screenSize = pag.size()
    currentItem = []#[x1,y1,x2,y2]
    searchItemsList = []
    kb = keyboard.Controller()

    # announce start
    print('start')

    # start listener
    with mouse.Listener(on_click=onClick) as listener:
        listener.join()