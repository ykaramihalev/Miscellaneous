import os
import pandas as pd
import pyautogui
import time


def rl360_di(number, name, amount, sedol, asset, ccy):
    os.startfile('Utmost International Dealing Instruction.pdf')
    time.sleep(10)

    for i in range(3):
        pyautogui.press('tab')

    pyautogui.write(name)

    for i in range(2):
        pyautogui.press('tab')

    pyautogui.write(number)

    for i in range(3):
        pyautogui.press('tab')

    pyautogui.write("Sell")
    pyautogui.press('tab')
    pyautogui.write(ccy)
    pyautogui.press('tab')
    pyautogui.write(amount)

    for i in range(3):
        pyautogui.press('tab')

    pyautogui.write(sedol)
    pyautogui.press('tab')
    pyautogui.write(asset)

    pyautogui.hotkey('ctrl', 'shift', 's')
    time.sleep(3)
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.write(number + ' ' + name)
    pyautogui.press('enter')


sales = pd.read_excel("Sales.xlsx")
sales["Value"] = sales["Value"].astype(str).copy()
sales["Policy Number"] = sales["Policy Number"].astype(str).copy()

for index, row in sales.iterrows():
    rl360_di(sales["Policy Number"][index], sales["Member Name"][index], sales["Value"][index], sales["ISIN"][index],
             sales["Asset Name"][index], sales["Currency"][index])

