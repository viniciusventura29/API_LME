from multiprocessing.managers import Namespace
import win32com.client as cli
from win32com.client import Dispatch

outlook = cli.Dispatch('Outlook.Application')
namespace = outlook.GetNamespace("MAPI")

acc = namespace.Folders['Campinas.ETS@br.bosch.com']

inbox = acc.Folders('Inbox')

# print(inbox.Items.Count)

from win32com.client import Dispatch

all_inbox = inbox.Items

for msg in all_inbox:
       if msg.Class==43:
           if msg.SenderEmailType=='EX':
               pass
           else:
               if msg.SenderEmailAddress == "viniciusventura29@icloud.com":
                    print(msg)
                    teste_folder = inbox.Folders.Add('Teste')

                    for message in all_inbox:
                        message.Move(teste_folder)