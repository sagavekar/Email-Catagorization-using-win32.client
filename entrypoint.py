import win32com.client
from icecream import ic
from notifypy import Notify
import sys 

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI") # Connect to Outlook
inbox = outlook.GetDefaultFolder(6)  # 6 corresponds to the inbox folder
messages = inbox.Items # Get all emails from the 
messages.Sort("[ReceivedTime]", True) # Sort emails by ReceivedTime, latest first

"""print(sys.getsizeof(messages))  --> to print whole size of inbox in bytes
message = inbox.Items.GetFirst() # Get the first email
message = inbox.Items.GetNext()# Get next email"""


def cat(emailcount:int)-> None:
    for index,message in enumerate(messages):
        if index==emailcount:
            break
        elif "sneha patra" in message.Body.lower():
            message.Categories = "sneha patra"
            message.Save()
        elif "kapil shakla" in message.Body.lower():
            message.Categories = "kapil shakla"
            message.Save()     
        elif "madhu gautam" in message.Body.lower():
            message.Categories = "madhu gautam"
            message.Save()
        elif "shantraj yeroor" in message.Body.lower():
            message.Categories = "shantraj yeroor"
            message.Save()
        elif "aditya ghadigaonkar" in message.Body.lower():
            message.Categories = "aditya ghadigaonkar"
            message.Save()    


        # Send desktop notification after completion of above code
        notification = Notify()
        notification.application_name = "Email categor Omkar.Sagavekar@gep.com"
        notification.title = ""
        notification.message = "Emails have been categorized"
        notification.send()    

if __name__ == '__main__':
    cat(emailcount=1)


