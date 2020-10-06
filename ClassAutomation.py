import subprocess
import os
import pyautogui
import getpass
import json
import datetime
import time
import win32com.client
from selenium import webdriver
from tkinter import *
from win10toast import ToastNotifier

notification = ToastNotifier()
name = getpass.getuser()
file_info = ""
date = datetime.datetime.now()
path = os.path.dirname(os.path.abspath(__file__))
webex = teams = zoom = False
root = webex_day = webex_link = webex_time = ""
day_list = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
teams_day = teams_link = teams_time = zoom_time = zoom_day = zoom_link = ""

def launch_webex(meeting=""):
    """If called, will launch Cisco Webex and join a certain meeting if was specified"""
    global file_info
    dumpling = file_info["information"][0]
    subprocess.Popen(dumpling["path"])
    time.sleep(5)
    try:
        meeting = meeting.split(" ")
        if meeting[2]:
            pyautogui.click(1603, 309, duration=0.25)
            time.sleep(2)
            pyautogui.typewrite(meeting[2])
            pyautogui.click(1833, 311, duration=0.25)
    except IndexError:
        pass
    time.sleep(56)

def launch_teams(meeting=""):
    """If called, will launch Microsoft Teams and join a certain meeting if was specified"""
    global file_info
    dumpling = file_info["information"][1]
    subprocess.Popen()
    time.sleep(5)
    pyautogui.click(65, 225, duration=0.25)
    pyautogui.click(65, 225, duration=0.25)
    time.sleep(56)

def launch_zoom(meeting=""):
    """If called, will launch Zoom and join a certain meeting if was specified"""
    global file_info
    dumpling = file_info["information"][2]
    subprocess.Popen()
    time.sleep(5)
    try:
        meeting = meeting.split(" ")
        if meeting[2]:
            pyautogui.click(970, 514, duration=0.25)
            time.sleep(2)
            pyautogui.click(907, 467, duration=0.25)
            pyautogui.typewrite(meeting[2])
    except IndexError:
        pass
    time.sleep(56)

def launch_meet(meeting=""):
    """Launches Google Chrome if available and enters Google Meet site"""
    driver = webdriver.Chrome()
    driver.get('https://meet.google.com/')
    time.sleep(5)
    try:
        meeting = meeting.split(" ")
        if meeting[2]:
            driver.find_element_by_xpath('//*[@id="page-content"]/section[1]/div/div[1]/div[2]/div/div[2]/input').send_keys(meeting[2])
            time.sleep(1)
            driver.find_element_by_xpath('//*[@id="page-content"]/section[1]/div/div[1]/div[2]/div/div[2]/a/button').click()
    except IndexError:
        pass
    time.sleep(56)

def find_webex():
    """Function that finds Cisco Webex Meetings executable"""
    global webex
    path1 = rf'C:\Users\{name}\AppData\Local\WebEx\WebEx\Applications\ptoneclk.exe'
    path2 = rf'C:\Program Files (x86)\Webex\Webex\Applications\ptoneclk.exe'
    if os.path.isfile(path1):
        webex = True
        return path1
    elif os.path.isfile(path2):
        webex = True
        return path2
    else:
        return False

def find_teams():
    """Function that finds Microsoft Teams executable"""
    global teams
    path0 = rf"C:\Users\{name}\AppData\Local\Microsoft\Teams\Update.exe"
    if os.path.isfile(path0):
        teams = True
        return path0
    else:
        return False

def find_zoom():
    """Function that finds Zoom executable"""
    global zoom
    path0 = rf"C:\Users\{name}\AppData\Roaming\Zoom\bin\Zoom.exe"
    if os.path.isfile(path0):
        zoom = True
        return path0
    else:
        return False

def find_browser():
    """Finds the default web browser in the users machine and if it is Google Chrome, return its path"""
    path0 = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    if os.path.isfile(path0):
        return path0
    else:
        return False

def ask_program():
    """Function will ask to the user which programs to open"""
    global file_info
    dumpling = file_info["information"]
    print("Input format: 'day' 'time' 'conference id'\nConference id is OPTIONAL")
    for i in dumpling:
        if i["path"]:
            selec = input(f"Hello {name}, you have {i['name']} available to execute, would you like to program a schedule for it? ")
            if selec == "yes" or selec == "y":
                i["needs_to_execute"] = 1
                while selec != "q":
                    selec = input(f"(Enter q anytime to quit)\nAt what day and time would you like to execute {i['name']}? ")
                    i['days'].append(selec)
                i["days"].pop()  

def init_file(w, t, z, m):
    """Does the initial setup for the configuration file"""
    global file_info
    file_info = {"information": [{"name": "Cisco Webex", "path": w, "needs_to_execute": 0, "days": []},
        {"name": "Microsoft Teams", "path": t, "needs_to_execute": 0, "days":[]},
        {"name": "Zoom", "path": z, "needs_to_execute": 0, "days":[]},
        {"name": "Google Meet", "path": m, "needs_to_execute": 0, "days":[]}]}

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

def launch_notification():
    """Prompts a desktop notification to the user that the program will execute in 5 minutes"""
    toaster.show_toast("ClassAutomation", "We will be joining your virtual class in 5 minutes!", icon_path="src/icon.ico", duration=5)

def main():
    """Main function of ClassAutomation"""
    global file_info, date

    dumpling = file_info["information"]
    day = date.strftime("%A").lower()
    times = date.strftime("%H") + ":" + date.strftime("%M")

    if times[0] == "0":
        times = times[1:]

    if int(times[-2:])-5 >= 0: minus5_min = int(times[-2:])-5
    else: minus5_min = 55
    
    for i in dumpling:
        if i["needs_to_execute"] == 1:
            dumpling2 = i["days"]
            for f in dumpling2:
                if day in f and times[:-2] + str(minus5_min) in date:
                    launch_notification()
                if day in f and times in f:
                    if "Webex" in i["name"]:
                        launch_webex(f)
                    elif "Teams" in i["name"]:
                        launch_teams(f)
                    elif "Zoom" in i["name"]:
                        launch_zoom(f)
                    elif "Meet" in i["name"]:
                        launch_meet(f)
                    
#START OF GUI->

def webex_getter():
    """Will add a new Cisco Webex schedule"""
    global webex_time, webex_day, webex_link, file_info
    if webex_day.get() and webex_time.get():
        dumpling = file_info["information"][0]
        dumpling["needs_to_execute"] = 1
        if webex_link.get():
            val = webex_day.get() + " " + webex_time.get() + " " + webex_link.get()
        else:
            val = webex_day.get() + " " + webex_time.get()
        dumpling["days"].append(val)
    else:
        pass

def teams_getter():
    """Will add a new Microsoft Teams schedule"""
    global teams_time, teams_day, teams_link, file_info
    if teams_day.get() and teams_time.get():
        dumpling = file_info["information"][1]
        dumpling["needs_to_execute"] = 1
        if teams_link.get():
            val = teams_day.get() + " " + teams_time.get() + " " + teams_link.get()
        else:
            val = teams_day.get() + " " + teams_time.get()
        dumpling["days"].append(val)
    else:
        pass

def zoom_getter():
    """Will add a new Zoom schedule"""
    global zoom_time, zoom_day, zoom_link, file_info
    if zoom_day.get() and zoom_time.get():
        dumpling = file_info["information"][2]
        dumpling["needs_to_execute"] = 1
        if zoom_link.get():
            val = zoom_day.get() + " " + zoom_time.get() + " " + zoom_link.get()
        else:
            val = zoom_day.get() + " " + zoom_time.get()
        dumpling["days"].append(val)
    else:
        pass

def gui_webex():
    """Enter schedule details if Cisco Webex is available"""
    global teams, zoom, webex_time, webex_day, webex_link, day_list
    webex_link = StringVar()
    webex_day = StringVar()
    webex_time = StringVar()
    Canvas(root, width=1000, height=1000).place(x=0, y=0)
    root.title("ClassAutomation: Cisco Webex")
    Label(root, text="Build your Cisco Webex schedule!", font=("Arial", 14)).place(x=110, y=1)
    Entry(root, textvariable=webex_link, width=20).place(x=350, y=200)
    Button(root, text="Submit", font=("Arial", 10), command=webex_getter).place(x=550, y=270)
    Entry(root, textvariable=webex_time, width=10).place(x=40, y=200)
    OptionMenu(root, webex_day, *day_list).place(x=200, y=195)
    Label(root, text="Time:", font=("Arial", 16)).place(x=40, y=150)
    Label(root, text="Day:", font=("Arial", 16)).place(x=200, y=150)
    Label(root, text="Session: (Optional)", font=("Arial", 16)).place(x=350, y=150)
    if teams == True:
        Button(root, text="Skip to Microsoft Teams", command=gui_teams, font=("Arial", 7)).place(x=470, y=5)
    elif zoom == True:
        Button(root, text="Skip to Zoom", command=gui_teams, font=("Arial", 7)).place(x=470, y=5)
    

def gui_teams():
    """Enter schedule details if Microsoft Teams is available"""
    global zoom, teams_day, teams_link, teams_time, day_list
    teams_day = StringVar()
    teams_time = StringVar()
    Canvas(root, width=1000, height=1000).place(x=0, y=0)
    Button(root, text="Submit", font=("Arial", 10), command=teams_getter).place(x=550, y=270)
    Entry(root, textvariable=teams_time, width=10).place(x=40, y=200)
    OptionMenu(root, teams_day, *day_list).place(x=200, y=195)
    root.title("ClassAutomation: Microsoft Teams")
    Label(root, text="Build your Microsoft Teams schedule!", font=("Arial", 14)).place(x=110, y=1)
    Label(root, text="Time:", font=("Arial", 16)).place(x=40, y=150)
    Label(root, text="Day:", font=("Arial", 16)).place(x=200, y=150)
    if zoom == True:
        Button(root, text="Skip to Zoom", command=gui_zoom, font=("Arial", 7)).place(x=500, y=5)

def gui_zoom():
    """Enter schedule details if Zoom is available"""
    global zoom_day, zoom_link, zoom_time, day_list
    zoom_link = StringVar()
    zoom_day = StringVar()
    zoom_time = StringVar()
    Canvas(root, width=1000, height=1000).place(x=0, y=0)
    Entry(root, textvariable=zoom_link, width=20).place(x=350, y=200)
    Button(root, text="Submit", font=("Arial", 10), command=zoom_getter).place(x=550, y=270)
    Entry(root, textvariable=zoom_time, width=10).place(x=40, y=200)
    OptionMenu(root, zoom_day, *day_list).place(x=200, y=195)
    Label(root, text="Time:", font=("Arial", 16)).place(x=40, y=150)
    Label(root, text="Day:", font=("Arial", 16)).place(x=200, y=150)
    Label(root, text="Session: (Optional)", font=("Arial", 16)).place(x=350, y=150)
    Button(root, text="Done", font=("Arial", 12), command= lambda: root.destroy()).place(x=500, y=5)
    root.title("ClassAutomation: Zoom")
    Label(root, text="Build your Zoom schedule!", font=("Arial", 14)).place(x=110, y=1)

def gui():
    """Main window to enter schedule details"""
    global root, webex, teams, zoom
    root = Tk()
    root.title("ClassAutomation")
    root.geometry("620x310")
    root.iconbitmap(path + r"\src\icon.ico")
    Label(root, text=f"Welcome to ClassAutomation {name}!", font=("Arial", 18)).place(x=50, y=0)
    if webex == True:
        Button(root, text="Create Cisco Webex schedule", command=gui_webex, font=("Arial", 14)).place(x=145, y=160)
    elif teams == True:
        Button(root, text="Create Microsoft Teams schedule", command=gui_teams, font=("Arial", 14)).place(x=140, y=160)
    elif zoom == True:
        Button(root, text="Create Zoom schedule", command=gui_zoom, font=("Arial", 14)).place(x=180, y=160)
    else:
        Label(root, text="It doesn't seem you have any compatible", font=("Arial", 18)).place(x=10, y=50)
        Label(root, text="applications in your system, try again later.", font=("Arial", 18)).place(x=10, y=100)
        Label(root, text="Oops!", font=("Arial", 42)).place(x=380, y=190)
    root.mainloop()

if os.path.isfile(r"C:\Users\Public\Documents\info.json"):
    with open(r"C:\Users\Public\Documents\info.json", "r") as info:
        file_info = json.load(info)
    while True:
        date = datetime.datetime.now()
        main()
        time.sleep(50)
else:
    webex_path = find_webex()
    teams_path = find_teams()
    zoom_path = find_zoom()
    meet_path = find_browser()
    init_file(webex_path, teams_path, zoom_path, meet_path)
    gui()
    add_to_startup_and_shortcut()
    with open(r"C:\Users\Public\Documents\info.json", "w") as info:
        json.dump(file_info, info)
    while True:
        date = datetime.datetime.now()
        main()
        time.sleep(50)