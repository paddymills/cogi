
from tqdm import tqdm
import time
import pyautogui

mouseloc = (148, 235)

docs = list()
with open('temp/docs.txt') as f:
    for row in f.readlines():
        row = row.strip()
        if row:
            docs.append(row)

for doc in tqdm(docs, desc="docs"):
    tqdm.write(f"Processing doc {doc}")

    pyautogui.click(*mouseloc)
    pyautogui.typewrite(doc)

    # if different year
    # pyautogui.hotkey('tab')
    # pyautogui.typewrite('2022')
    
    pyautogui.hotkey('enter')
    time.sleep(0.5)

    pyautogui.hotkey('ctrl', 's')
    time.sleep(0.5)

    # pyautogui.hotkey('f12')
    # pyautogui.hotkey('f12')
    # pyautogui.hotkey('f12')
    # pyautogui.click(118, 539)
    # pyautogui.moveTo(*mouseloc)
    # pyautogui.hotkey('enter')


    x, y = pyautogui.position()
    if mouseloc != (x, y):
        break