import os
import pandas as pd
import pyautogui
import time

os.chdir(r'M:\Investments\Automation\Python\Dealing Instructions')


def rl360_di(number, name, amount, sedol, asset, ccy):
    """Completes RL360 dealing form."""

    os.startfile('RL360 Blank Dealing.pdf')
    time.sleep(10)

    for i in range(2):
        pyautogui.press('tab')

    pyautogui.write(number)
    pyautogui.press('tab')
    pyautogui.write(name)

    for i in range(9):
        pyautogui.press('tab')

    pyautogui.write(amount)

    for i in range(2):
        pyautogui.press('tab')

    pyautogui.write(sedol)
    pyautogui.press('tab')
    pyautogui.write(asset)
    pyautogui.press('tab')
    pyautogui.write(ccy)

    pyautogui.hotkey('ctrl', 'shift', 's')  # file is being saved in default Documents folder
    time.sleep(3)
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.write(number + ' ' + name)
    pyautogui.press('enter')


sales = pd.read_excel("Sales.xlsx")
sales["Value"] = sales["Value"].astype(str).copy()  # ensuring correct datatypes
sales["Policy Number"] = sales["Policy Number"].astype(str).copy()


for index, row in sales.iterrows():  # iterating over rows in file and using the values as inputs in the rl360 function
    rl360_di(sales["Policy Number"][index], sales["Member Name"][index], sales["Value"][index], sales["ISIN"][index],
             sales["Asset Name"][index], sales["Currency"][index])



