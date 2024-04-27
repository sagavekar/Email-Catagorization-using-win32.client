import win32com.client
from icecream import ic


# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Define the name of the subfolder
subfolder_name = "ABM Catalog"

# Function to recursively search for a subfolder by name

def find_subfolder(folder, name):
    if folder.Name == name:
        return folder
    else:
        for subfolder in folder.Folders:
            result = find_subfolder(subfolder, name)
            if result is not None:
                return result

# Start the search from the root folder
root_folder = outlook.Folders.Item(1) # omkar.sagavekar@gep.com
root_folder = outlook.Folders.Item(2) # SET interface
inbox = outlook.GetDefaultFolder(6) # inbox

inbox_emails = inbox.Items
inbox_emails.Sort("[ReceivedTime]", True) # Sort emails by ReceivedTime, latest first
match = "Omkar Sagavekar"
for index,message in enumerate(inbox_emails):
    if match in message.Body:
        message.Categories = "Omkar Sagavekar"
        message.Save()
    if index==20:
        break


# subfolder = find_subfolder(root_folder, subfolder_name)

# # Check if subfolder was found
# if subfolder:
#     print("Subfolder found:", subfolder.Name)
# else:
#     print("Subfolder not found.")