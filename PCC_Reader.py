#=================================================================================
#
#                            © Criado por Raphael Frei
#                                2022 - Adient PLC    
#
#=================================================================================

# Run via CMD by the following command:
# "file_path\PCC_Reader.exe" "email_to_check" "file_path\save_location"

# Works with Windows Language pt_BR or en_US

import sys
import os
import ctypes
import locale

import win32com.client

from datetime import datetime, timedelta

windll = ctypes.windll.kernel32
language = locale.windows_locale[windll.GetUserDefaultUILanguage()]

outlook = win32com.client.Dispatch("Outlook.Application")

mapi = outlook.GetNamespace("MAPI")

try:
    inbox = mapi.Folders(sys.argv[1]).Folders('Caixa de Entrada')    
except:
    try:
        inbox = mapi.Folders(sys.argv[1]).Folders('Inbox')
    except Exception as e:
        print("Error when selection folders: " + str(e))

messages = inbox.Items

if(language == "pt_BR"):
    messages = messages.Restrict("[ReceivedTime] >= '" + (datetime.now() - timedelta(days = 1)).strftime('%d/%m/%Y %H:%M %p') + "'")
elif(language == "en_US"):
    messages = messages.Restrict("[ReceivedTime] >= '" + (datetime.now() - timedelta(days = 1)).strftime('%m/%d/%Y %H:%M %p') + "'")

try:
    for message in list(messages):
        if message.unread == True:
            if not os.path.exists(sys.argv[2]):
                os.makedirs(sys.argv[2])
            message.unread = False
            message.SaveAs(os.path.join(sys.argv[2], (message.Subject + ".msg")))
except Exception as e:
    print("Error when processing emails: " + str(e))