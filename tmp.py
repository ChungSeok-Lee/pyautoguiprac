import pywinauto
from pywinauto import application
import time
import pyautogui
import pandas as pd 
from EditName import EditName
import PIL
from PIL import ImageGrab


# app = application.Application().start('C:/Program Files (x86)/GOM/GOMMix/GomMixMain.exe')
# time.sleep(5)
# pyautogui.hotkey("Alt", "F4")
# app.connect(process = 12844)
# app.


# main_dlg = app.window(title = 'Untitled')
# main_dlg.print_control_identifiers()

# df = pd.read_excel('./File/1차_합산본_20201215.xlsx')
# df = df.dropna()
# print(df.tail())

# for i in range(31):
#     tno = i % 10
#     if tno == 0:
#         print(i, "True")
#     else:
#         if (i+1) % 10 ==0:
#             print(i, "Pre-True")
        
#         else:
#             print(i, "False")

print(pywinauto.__version__)
print(pyautogui.__version__)
print(pd.__version__)
print(PIL.__version__)
