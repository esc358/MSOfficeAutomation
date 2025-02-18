import win32com.client
import csv
from datetime import datetime

#Instantiate Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

#Access calendar folder
calendar = outlook.GetDefaultFolder(9)

print(calendar)




