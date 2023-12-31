import os
import sys
import pandas as pd
import pyautogui
import time

os.chdir(r'M:\Investments\Automation\Python\Dealing Instructions')


def utmost_di(number, name, amount, sedol, asset, ccy):
    """Completes Utmost International dealing form."""

    try:
        os.startfile('Utmost International Dealing Instruction.pdf')
    except FileNotFoundError:
        print("\nThe process failed. Please ensure the Utmost International dealing template is in the folder before "
              "running the program again.")
        sys.exit()

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

    pyautogui.hotkey('ctrl', 'shift', 's')  # file is being saved in default Documents folder
    time.sleep(3)
    pyautogui.press('enter')
    time.sleep(3)
    pyautogui.write(number + ' ' + name)
    pyautogui.press('enter')


try:
    sales = pd.read_excel("Sales.xlsx")
except FileNotFoundError:
    print("\nThe process failed. Please ensure the Sales file is in the folder before running the program again.")
    sys.exit()

sales["Value"] = sales["Value"].astype(str).copy()
sales["Policy Number"] = sales["Policy Number"].astype(str).copy()

for index, row in sales.iterrows():  # iterating over rows in file and using the values as inputs in the utmost function
    utmost_di(sales["Policy Number"][index], sales["Member Name"][index], sales["Value"][index], sales["ISIN"][index],
             sales["Asset Name"][index], sales["Currency"][index])

