from pynput.keyboard import *
from datetime import date
import pyautogui
import time

today = date.today().strftime("%m/%d/%Y")
orders = [
    "1000890217",
    "1000890213",
    "1000890219",
    "1000890220",
    "1000888697",
    "1000887869",
    "1000887848",
    "1000887850",
    "1000887851",
    "1000887854",
    "1000887856",
    "1000887859",
    "1000887861",
    "1000887874",
    "1000887876",
    "1000887860",
    "1000887862",
    "1000887864",
    "1000888709",
    "1000888776",
    "1000888780",
    "1000899931",
    "1000887877",
    "1000887879",
    "1000887880",
    "1000887882",
    "1000887871",
    "1000887873",
    "1000887875",
    "1000887881",
    "1000888800",
    "1000888808",
    "1000888811",
]


def on_release(key):
    if key == KeyCode.from_char("a"):
        pyautogui.press("backspace")
        pyautogui.typewrite(orders.pop())
        pyautogui.press("enter")

    elif key == KeyCode.from_char("d"):
        pyautogui.press("backspace")
        pyautogui.typewrite(today)
        pyautogui.hotkey("ctrl", "s")
        time.sleep(0.5)
        pyautogui.press("f3")

    if key == Key.esc or key == KeyCode.from_char("x"):
        # stop listener and exit
        return False

    if len(orders) == 0:
        print("All orders processed")
        return False


with Listener(on_release=on_release) as listener:
    print("ready...")
    listener.join()
