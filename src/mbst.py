from tqdm import tqdm
import time
import pyautogui

user_dwell = 3

docs = list()
with open("temp/docs.txt") as f:
    for row in f.readlines():
        row = row.strip()
        if row:
            docs.append(row)


print("Place your cursor at the start of the `Material Doc` field")

positions = []
frame_dur = 0.25
while 1:
    positions.append(pyautogui.position())
    time.sleep(frame_dur)

    # mouse has not moved in {user_dwell} seconds
    if len(positions) > user_dwell * frame_dur and len(set(positions[-8:])) == 1:
        break

mouseloc = positions[-1]
pyautogui.click(*mouseloc)

for doc in tqdm(docs, desc="docs"):
    tqdm.write(f"Processing doc {doc}")

    pyautogui.click(*mouseloc)
    pyautogui.press("delete", presses=10)
    pyautogui.typewrite(doc)

    # if different year
    pyautogui.hotkey("tab")
    # pyautogui.typewrite('2022')
    pyautogui.press("delete", presses=4)

    pyautogui.press("enter")
    time.sleep(0.5)

    pyautogui.hotkey("ctrl", "s")
    time.sleep(0.5)

    # exit to main and re-enter
    # pyautogui.press('f12', presses=3)
    # pyautogui.click(118, 539)
    # pyautogui.moveTo(*mouseloc)
    # pyautogui.press('enter')

    x, y = pyautogui.position()
    if mouseloc != (x, y):
        break
