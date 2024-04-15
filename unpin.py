import win32com.client
import time
from datetime import datetime
import csv

time.sleep(120)

logs_path = "C:/Users/viacheslav.pustovit/Downloads/MIS PROYECTOS/unpin_logs.csv"

def stamp():
    with open(logs_path, "a") as f:
        csv.writer(f).writerow([datetime.now()])

def unpin_from_taskbar(app_name):
    shell = win32com.client.Dispatch("Shell.Application")
    # Get the Taskbar object
    taskbar = shell.Namespace("C:\ProgramData\Microsoft\Windows\Start Menu\Programs")
    # Find the application icon in the Taskbar
    for i in range(taskbar.Items().Count):
        item = taskbar.Items().Item(i)
        #print(item)  # if you wnat see the available files in the file
        if item.Name == app_name:
            # Unpin the application from the Taskbar
            verb = None
            for v in item.Verbs():
                if "Unpin from tas&kbar" in str(v):
                    verb = v
                    break
            if verb:
                verb.DoIt()
                print(f"{app_name} unpinned from the taskbar.")
            else:
                print(f"Could not find 'Unpin from taskbar' verb for {app_name}.")
            break
    else:
        print(taskbar)
        print(f"{app_name} is not pinned to the taskbar.")


def print_available_verbs(file_path):
    shell = win32com.client.Dispatch("Shell.Application")
    folder_item = shell.Namespace(0).ParseName(file_path)
    if folder_item is not None:
        verbs = [verb.Name for verb in folder_item.Verbs()]
        print("Available Verbs:")
        for verb in verbs:
            print(verb)
    else:
        print("File not found.")

# Specify the name of the application you want to unpin from the taskbar
unpin_from_taskbar("Excel")
unpin_from_taskbar("Word")
unpin_from_taskbar("Outlook")

stamp()

#check the available verbs (options when you perform a right click mouse, Open a file, Edit file, properties, Pin to taskbar)
#file_path = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Word.lnk"  # Replace with the path to your file
#print_available_verbs(file_path)
