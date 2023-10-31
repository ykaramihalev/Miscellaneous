import pandas as pd
import win32com.client as win32
import os

os.chdir(r"M:\Investments\Automation\Python\Email")

send_list = pd.read_excel("Adviser Emails.xlsx")
advisers = send_list["Adviser Firm"].drop_duplicates()

for adviser_firm in advisers:

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Subject = f"Overdrafts {adviser_firm}"
    recip = mail.Recipients.Add('dealing@ipensionsgroup.com')
    recip.Type = 2

    send_list_new = send_list[send_list["Adviser Firm"] == adviser_firm]
    mail.To = send_list_new["Adviser Email Address"].tolist()[0]

    mail.GetInspector

    message = f"""Dear Team,<p>I am writing to you with regards to the investment cash account(s) for some of your 
    clients with us. The accounts are currently showing as overdrawn and require your immediate action to restore the 
    cash balance to a positive position. This can be done by submitting a dealing instruction to us. Below are the 
    details of the current cash positions on the accounts:<p>
    {send_list_new[["Policy Number", "Member Name", "Investment Company", "Overdraft Value", "Currency"]].to_html(index=False)}<p>
    When submitting the dealing instruction, please also ensure it provides for sufficient cash available for any future
    fees and charges payable in upcoming months, to avoid the account reverting to an overdrawn position and any 
    unnecessary interest payments for the member. We also ask you to review any upcoming benefit payments requested by 
    the member, prior to submitting the dealing instruction.<p>Please note the deadline to complete this request is 
    14/08/2023, if we do not receive a dealing instruction by the deadline, we will proceed with selling from the 
    client’s highest most liquid asset. If the overdraft has been cleared, please disregard this email."""

    index = mail.HTMLbody.find('>', mail.HTMLbody.find('<body'))
    mail.HTMLbody = mail.HTMLbody[:index + 1] + message + mail.HTMLbody[index + 1:]
    mail.Display()
    mail.Send()