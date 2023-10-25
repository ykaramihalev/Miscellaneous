import os
import pandas as pd
import pyautogui
import time

os.chdir(r'M:\Investments\Automation\Python\Dealing Instructions')


def fpi_di(number, name, amount, sedol, asset, ccy):
    """Completes Friends Provident International dealing form."""

    os.startfile('FPI Dealing Instruction.pdf')
    time.sleep(10)

    for i in range(7):
        pyautogui.press('tab')

    pyautogui.write(name)
    pyautogui.press('tab')
    time.sleep(2)
    pyautogui.write(number)
    pyautogui.press('tab')
    pyautogui.write('Sell')
    pyautogui.press('tab')
    pyautogui.write(ccy)
    pyautogui.press('tab')
    pyautogui.write(amount)

    for i in range(2):
        pyautogui.press('tab')

    pyautogui.write(sedol)
    pyautogui.press('tab')
    pyautogui.write(asset)
    pyautogui.press('tab')
    pyautogui.write(ccy)

    pyautogui.hotkey('ctrl', 'shift', 's')   # file is being saved in default Documents folder
    time.sleep(3)
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.write(number + ' ' + name)
    pyautogui.press('enter')


sales = pd.read_excel("Sales.xlsx")
sales["Value"] = sales["Value"].astype(str).copy()
sales["Policy Number"] = sales["Policy Number"].astype(str).copy()

for index, row in sales.iterrows():  # iterating over rows in file and using the values as inputs in the fpi function
    fpi_di(sales["Policy Number"][index], sales["Member Name"][index], sales["Value"][index], sales["ISIN"][index],
             sales["Asset Name"][index], sales["Currency"][index])

