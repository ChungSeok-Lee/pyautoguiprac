import pyautogui as pg
import time
from PIL import ImageGrab


# screen = ImageGrab.grab()
# print(screen.getpixel((1131, 519)))

while True:
    screen = ImageGrab.grab()
    print(screen.getpixel((1476, 518)))
    if screen.getpixel((1476, 518)) == (217, 52, 66):
        print('END')
        break
    time.sleep(1)