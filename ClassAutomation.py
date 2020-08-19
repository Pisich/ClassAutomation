import subprocess
import os
import pyautogui
import getpass
import json
import datetime
import time
import win32com.client

name = getpass.getuser()
file_info = ""
date = datetime.datetime.now()

def find_webex():
    """Function that finds Cisco Webex Meetings executable"""
    path1 = rf'C:\Users\{name}\AppData\Local\WebEx\WebEx\Applications\ptoneclk.exe'
    path2 = rf'C:\Program Files (x86)\Webex\Webex\Applications\ptoneclk.exe'
    if os.path.isfile(path1):
        return path1
    elif os.path.isfile(path2):
        return path2
    else:
        return False

def find_teams():
    """Function that finds Microsoft Teams executable"""
    path0 = rf"C:\Users\{name}\AppData\Local\Microsoft\Teams\Update.exe"
    if os.path.isfile(path0):
        return path0
    else:
        return False

def find_zoom():
    """Function that finds Zoom executable"""
    path0 = rf"C:\Users\{name}\AppData\Roaming\Zoom\bin\Zoom.exe"
    if os.path.isfile(path0):
        return path0
    else:
        return False

def ask_program():
    """Function will ask to the user which programs to open"""
    global file_info
    dumpling = file_info["information"]
    for i in dumpling:
        if i["path"]:
            selec = input(f"Hello {name}, you have {i['name']} available to execute, would you like to program a schedule for it? ")
            if selec == "yes" or selec == "y":
                i["needs_to_execute"] = 1
                while selec != "q":
                    selec = input(f"(Enter q anytime to quit)\nAt what day and time would you like to execute {i['name']}? ")
                    i['days'].append(selec)
                i["days"].pop()  

def init_file(w, t, z):
    """Does the initial setup for the configuration file"""
    global file_info
    file_info = {"information": [{"name": "Cisco Webex", "path": w, "needs_to_execute": 0, "days": []},
        {"name": "Microsoft Teams", "path": t, "needs_to_execute": 0, "days":[]},
        {"name": "Zoom", "path": z, "needs_to_execute": 0, "days":[]}]}

def add_to_startup_and_shortcut(file_path=""):
    if file_path == "":
        file_path = os.path.dirname(os.path.realpath(__file__))
    file_name = r"\ClassAutomation"
    bat_path = r'C:\Users\%s\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup' % name
    with open(bat_path + '\\' + "open.bat", "w+") as bat_file:
        bat_file.write(r'start "" C:/Users/Public/Documents%s.lnk' % file_name)
    desktop = r'C:/Users/Public/Documents'
    path = os.path.join(desktop, 'ClassAutomation.lnk')
    target = file_path + file_name + ".exe"

    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.WindowStyle = 7
    shortcut.save()

def main():
    """Main function of ClassAutomation"""
    global file_info, date
    dumpling = file_info["information"]
    day = date.strftime("%A").lower()
    times = date.strftime("%H") + ":" + date.strftime("%M")
    if times[0] == "0":
        times = times[1:]
    for i in dumpling:
        if i["needs_to_execute"] == 1:
            dumpling = i["days"]
            for f in dumpling:
                if day in f and times in f:
                    subprocess.Popen(i["path"])
                    time.sleep(61)


if os.path.isfile(r"C:\Users\Public\Documents\info.json"):
    with open(r"C:\Users\Public\Documents\info.json", "w") as info:
        file_info = json.load(info)
    while True:
        date = datetime.datetime.now()
        main()
else:
    webex_path = find_webex()
    teams_path = find_teams()
    zoom_path = find_zoom()
    init_file(webex_path, teams_path, zoom_path)
    ask_program()
    add_to_startup_and_shortcut()
    while True:
        date = datetime.datetime.now()
        main()

with open(r"C:\Users\Public\Documents\info.json", "w") as info:
    json.dump(file_info, info)
