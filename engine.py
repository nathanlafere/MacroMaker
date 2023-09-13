import keyboard
import pyautogui
import time
import os
import pickle

def run_macro(event_list,loops):
    counter = 0
    while counter > loops:
        counter += 1
        for event in event_list:
            match event[0]:
                case "Mouse left click":
                    time.sleep(event[2]/1000)
                    pyautogui.click(event[1][0]+counter*event[1][2],event[1][1]+counter*event[1][3])
                case "Mouse right click":
                    time.sleep(event[2]/1000)
                    pyautogui.click(event[1][0]+counter*event[1][2],event[1][1]+counter*event[1][3],button='secondary')
                case "Press key":
                    print(event[0])
                case "Copy":
                    print(event[0])
                case "Insert":
                    print(event[0])