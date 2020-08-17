import subprocess
import os
import pyautogui
import getpass


def start_webex(path):
    subprocess.Popen(path)


def main():
    name=getpass.getuser()
    path1=rf'C:\Users\{name}\AppData\Local\WebEx\WebEx\Applications\ptoneclk.exe'
    path2=rf'C:\Program Files (x86)\Webex\Webex\Applications\ptoneclk.exe'
    n=os.path.isfile(path1)
    n2=os.path.isfile(path2)
    if n== True:
        start_webex(path1)
    elif n2==True:
        start_webex(path2)



main()
