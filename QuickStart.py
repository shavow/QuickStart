
#███████╗██╗  ██╗ █████╗ ██╗   ██╗ ██████╗ ██╗    ██╗
#██╔════╝██║  ██║██╔══██╗██║   ██║██╔═══██╗██║    ██║
#███████╗███████║███████║██║   ██║██║   ██║██║ █╗ ██║
#╚════██║██╔══██║██╔══██║╚██╗ ██╔╝██║   ██║██║███╗██║
#███████║██║  ██║██║  ██║ ╚████╔╝ ╚██████╔╝╚███╔███╔╝
#╚══════╝╚═╝  ╚═╝╚═╝  ╚═╝  ╚═══╝   ╚═════╝  ╚══╝╚══╝ 

from math import *
import sys
from os import R_OK, W_OK, X_OK, close, mkdir, name, path, read, rename, spawnl, stat, times, listdir
from os.path import isfile, join
import os
import random
import time
import threading
from psutil import process_iter
import pyperclip
from threading import Thread, Timer, active_count, local, Event
from multiprocessing import Process, Value
import shutil
from shutil import copy2, register_archive_format
from pathlib import Path
import pathlib
from datetime import date, datetime
import os, subprocess
from tkinter import *
from tkinter import font
from tkinter.scrolledtext import ScrolledText
from tkinter import messagebox
from tkinter import filedialog
from tkinter import scrolledtext
import PIL
from PIL import ImageTk, Image, ImageColor
import win32com.client, pythoncom, winshell, psutil, pywinauto

#________________________| locals |________________________#

#ui locals --
lastClickX = 0
lastClickY = 0

inverted = "black"
inverted1 = "white"
inverted2 = "#de70ff"

g_fileNamesN = []
g_fileNames = []
g_filePaths = []

g_activePrograms = []
g_activeProgramsN = []

g_focusedTab = "openAppilcation" #this determines which tab is open currently, if "closeApplication" is the name of g_focusedTab then the closeApplication listbox be visible .

#other locals --
userPath = os.path.expanduser("~")
userName = userPath.split("\\")[-1]
roamingPath = (userPath + "\\AppData\\Roaming\\")

folderPath = (roamingPath + "QuickLOAD\\")
configFolder = (folderPath + "shortcuts\\")

#________________________| events and functions |________________________#

#automated commands and additional commands --
def command_recovery():
    if os.path.isdir(roamingPath):
        if os.path.isdir(folderPath):
            pass
        else:
            os.mkdir(folderPath) #make main folder .
        
        if os.path.isdir(configFolder):
            pass
        else:
            os.mkdir(configFolder) #make configs folder .

def colorConvert(r,g,b): #color converter used to make RGB code to HEX code .
    return f'#{r:02x}{g:02x}{b:02x}'

def detailedShavow(): #this command makes all the ui widgets go colourful .
    while True:
        if "255" in str(ImageColor.getcolor(label_creditCreator.cget("fg"), "RGB")).split(",")[1]: #gets the rgb code and converts to hex, 
            reversedRange = range(255, 0, -5)
            for number in reversedRange: 
                label_creditCreator.config(fg=colorConvert(255,number-5,255)) #then converts the oposite way to change color from pink-white
                button_openFolder.config(fg=colorConvert(255,number-5,255))
                button_closeWindow.config(fg=colorConvert(255,number-5,255))
                button_switchTabs.config(fg=colorConvert(255,number-5,255))
                listbox_shortcutHolder.config(selectforeground=colorConvert(255,number-5,255))
                time.sleep(.005)
        elif "0" in str(ImageColor.getcolor(label_creditCreator.cget("fg"), "RGB")).split(",")[1]:
            for x in range(52):
                label_creditCreator.config(fg=colorConvert(255,x*5,255)) #then converts the oposite way to change color from white-pink
                button_openFolder.config(fg=colorConvert(255,x*5,255))
                button_closeWindow.config(fg=colorConvert(255,x*5,255))
                button_switchTabs.config(fg=colorConvert(255,x*5,255))
                listbox_shortcutHolder.config(selectforeground=colorConvert(255,x*5,255))
                time.sleep(.005)

def checkOnTop(): #checks if the appilcation is ontop.checked = True,False and then applies it to the program .
    while True:
        time.sleep(.05)
        Value = variable.get()
        if Value == 1:
            window.attributes('-topmost', True)
        elif Value == 0:
            window.attributes('-topmost', False)

def autoUpdating(): #does the same as the following comment below, but instead it checks if a "new" file has been placed inside the folder directory .
    while True:
        time.sleep(0.5)
        files = os.scandir(configFolder)
        for f in files:
            if os.path.abspath(f) not in g_filePaths:
                refreshFolder()
        files.close()

def autoUpdatingCLOSE(): #this auto-updating thing will check if a file has been deleted in the folder directory, and if so, then refresh the listbox .
    while True:
        time.sleep(0.5)
        for f in g_filePaths:
            if "FILLAR" in f:
                pass
            elif os.path.isfile(f):
                pass
            else:
                refreshFolder()

def refreshFolder(): #this refresh command will completely wipe the already stored filepaths .
    global g_filePaths, g_fileNames, g_fileNamesN
    g_filePaths.clear()
    g_fileNames.clear()
    g_fileNamesN.clear()
    listbox_shortcutHolder.delete(0,"end") #and when a new file is placed in the folder directory .

    xTimes = 0
    folderScan = os.scandir(configFolder) #it will placed and added onto the listbox .
    for x in folderScan:
        fullPath = os.path.abspath(x) #adding the file to 3 different lists, name_of_file, name_of_file (with number to index filepath), file_path_of_file .
        if os.path.isfile(fullPath):
            xTimes +=1 #the number displays the index number for the file path .
            xName = fullPath.split("\\")[-1]
            xNameN = str(xTimes) + ". " + os.path.splitext(xName)[0] #pretty self-explanitory .
            xPath = fullPath

            g_fileNamesN.append(xNameN)
            g_fileNames.append(xName)
            g_filePaths.append(xPath)

            listbox_shortcutHolder.insert(END, xNameN)
    folderScan.close()

def insertAllPrograms(): #this will insert all processes and refresh the lists
    global g_activePrograms, g_activeProgramsN
    g_activeProgramsN.clear()
    g_activePrograms.clear()
    listbox_closeApplication.delete(0,"end")
    xTimes = 0

    for process in psutil.process_iter():

        processName = process.name()

        if "svchost.exe" in processName:
            pass
        else:
            if processName in g_activePrograms:
                pass
            else:
                xTimes +=1
                procName = (str(xTimes) + ". " + processName)
                g_activeProgramsN.append(procName)
                g_activePrograms.append(processName)
                listbox_closeApplication.insert("end", procName)

def createShortcut(filePath, fileName):
    FilePath1 = configFolder + fileName + ".lnk" #folderpath + namefile + .Ink [shortcut extension] .
    OpenPath = filePath
    Icon = filePath

    shortcut = win32com.client.Dispatch("WScript.Shell")
    CreateFile = shortcut.CreateShortCut(FilePath1) #where the shortcut will be placed (in the folder directory) .
    CreateFile.Targetpath = OpenPath #obviously the target once the shortcut is opened .
    CreateFile.IconLocation = Icon #uses the application that is wants to copy's icon .
    CreateFile.save() #creates shortcut to folder .

def findFilePath(file):
    if "openAppilcation" in g_focusedTab:
        split = file.split(". ")[0] #this will get the number displayed in the listbox and index it in the stored file paths .
        filePathed = g_filePaths[int(split)-1]
        return filePathed #then returning the file path to be used and collected .
    elif "closeApplication" in g_focusedTab:
        split = file.split(". ")[0]
        processName = g_activePrograms[int(split)-1]
        return processName

#ui functions -- 
def SaveLastClickPos(event): #this will make it so it will be possible to move the ui by holding down click .
    global lastClickX, lastClickY
    lastClickX = event.x
    lastClickY = event.y

def Dragging(event):
    x, y = event.x - lastClickX + window.winfo_x(), event.y - lastClickY + window.winfo_y() #an add on ^
    window.geometry("+%s+%s" % (x , y))

def command_closeWindow(): #this will quit the program .
    window.quit()
    os._exit(0)

def command_openFolder(): #this will open the folder directory .
    if os.path.isdir(folderPath):
        subprocess.Popen(f'explorer "/select", {folderPath}')
    else:
        command_recovery()
        subprocess.Popen(f'explorer "/select", {folderPath}') #either way it will recover the folder directory and open the folder path .

def command_importFileShortcut(): #this will import a exe file that you had selected from file explorer .
    global g_ifChanged
    if "openAppilcation" in g_focusedTab: #this will block it from adding exe files into the folder directory while being on the wrong tab .

        selectFile =filedialog.askopenfilename(initialdir=f"{userPath}\\Documents", title="SELECT AN EXE",
                    filetypes=(("EXE FILES","*.exe"), ("All Files", "*.*")))

        remarkedFile = selectFile.replace("/", "\\") #a readable format .

        if selectFile:
            fileName = remarkedFile.split("\\")[-1] #name of file .
            fileName = os.path.splitext(fileName) #splits name and extension .
            if len(fileName) > 1: #if has extension .
                if "EXE" in fileName[1].upper(): #if it is an exe .
                    filePath = configFolder + fileName[0]
                    if filePath not in g_filePaths: #if the filePath isnt already in the listbox .
                        createShortcut(remarkedFile, fileName[0]) #create the shortcut obviously .
        refreshFolder()
    elif "closeApplication" in g_focusedTab:
        insertAllPrograms()

def command_executeFile(): #this will execute the selected file and run the program .
    if "openAppilcation" in g_focusedTab:
        getSelected = str(listbox_shortcutHolder.get(ACTIVE))
        filePath = findFilePath(getSelected) #converts the NameN into the file path .
        if os.path.isfile(filePath):
            os.startfile(filePath) #this will start the file obviously .

    elif "closeApplication" in g_focusedTab:
        getSelected = str(listbox_closeApplication.get(ACTIVE))
        processName = findFilePath(getSelected)
        for process in psutil.process_iter():
            if processName in process.name():
                process.kill()
                insertAllPrograms()

def command_deleteFile(): #this will delete the selected file .
    if "openAppilcation" in g_focusedTab: #this will block it from accidently removing applications LOL
        getSelected = str(listbox_shortcutHolder.get(ACTIVE))
        filePath = findFilePath(getSelected) #converts the NameN into the file path .
        if os.path.isfile(filePath):
            os.remove(filePath) #this will remove the shortcut file .
            refreshFolder() #and then refresh the listbox .

def command_switchTabs():
    global g_focusedTab
    if ">" in button_switchTabs.cget("text"):
        button_switchTabs.config(text="<") #switch to close programs tab
        listbox_shortcutHolder.place_forget()
        listbox_closeApplication.place(x=32,y=5)

        g_focusedTab = "closeApplication"

    elif "<" in button_switchTabs.cget("text"):
        button_switchTabs.config(text=">") #switch to open programs tab
        listbox_closeApplication.place_forget()
        listbox_shortcutHolder.place(x=32,y=5)

        g_focusedTab = "openAppilcation"

#________________________| ui window |________________________#

#main window -- 
window = Tk()
window.title("QUICKStart")

window.geometry("500x150+540+200")
window.minsize(500,150)
window.maxsize(500,150)
window.config(background="#050505")
window.overrideredirect(True)
window.attributes('-topmost', False)
variable = IntVar(value=1)

#main canvas --
mainBorder =Frame(window, highlightbackground=inverted1, highlightthickness=1, bd=0)
main =Frame(mainBorder, bg=inverted, height=150, width=500, bd=0)

#main window1 [label_credit, button_ontop, button_close, button_import, button_execute, button_openfolder, button_delete] --
frame_lineFrame =Frame(main, highlightbackground=inverted1, highlightthickness=1, bd=0, width=1, height=150)

label_creditCreator =Label(main, text="S\nH\nA\nV\nV\nO\nW", font=("Name Smile", 14), bg=inverted, fg=inverted1, bd=0, relief=FLAT)

button_closeWindow =Button(main, relief=FLAT, text="X", font=("Name Smile", 12), bg=inverted, fg=inverted1, activebackground=inverted, activeforeground=inverted1, bd=0,
                            command =command_closeWindow)

button_openFolder =Button(main, relief=FLAT, text="0", font=("Name Smile", 12), bg=inverted, fg=inverted1, activebackground=inverted, activeforeground=inverted1, bd=0,
                            command =command_openFolder)

button_switchTabs =Button(main, relief=FLAT, text=">", font=("Name Smile", 12), bg=inverted, fg=inverted1, activebackground=inverted, activeforeground=inverted1, bd=0,
                            command =command_switchTabs)

listbox_shortcutHolder = Listbox(window, bg=inverted,  fg=inverted1, height=7, width=33, bd=0, border=0, borderwidth=0, highlightbackground=inverted1, highlightthickness=0,
                     font=("Name Smile", 11), selectbackground=inverted, selectforeground=inverted2)

listbox_closeApplication = Listbox(window, bg=inverted,  fg=inverted1, height=7, width=33, bd=0, border=0, borderwidth=0, highlightbackground=inverted1, highlightthickness=0,
                     font=("Name Smile", 11), selectbackground=inverted, selectforeground=inverted2)

button_importFile = Button(main, relief=FLAT, text="[  +  ]", font=("Name Smile", 10), bg=inverted, fg=inverted1, activebackground=inverted, activeforeground=inverted1, bd=0,
                            command =command_importFileShortcut)

button_executeFile = Button(main, relief=FLAT, text="[  !  ]", font=("Name Smile", 10), bg=inverted, fg=inverted1, activebackground=inverted, activeforeground=inverted1, bd=0,
                            command =command_executeFile)

button_deleteFile = Button(main, relief=FLAT, text="[  -  ]", font=("Name Smile", 10), bg=inverted, fg=inverted1, activebackground=inverted, activeforeground=inverted1, bd=0,
                            command =command_deleteFile)

button_alwaysOnTop =Checkbutton(main, relief=FLAT, text="ON TOP", font=("Name Smile", 10), bg=inverted, fg=inverted1, activebackground=inverted, activeforeground=inverted1,
                                selectcolor=inverted, variable=variable,)

#packs and places --
main.pack()
mainBorder.pack()

frame_lineFrame.place(x=25.5,y=0)
label_creditCreator.place(x=3.25, y=3)

button_closeWindow.place(x=470, y=125)
button_openFolder.place(x=445, y=125)
button_switchTabs.place(x=425, y=125)

button_importFile.place(x=90+20, y=128)
button_executeFile.place(x=30+20, y=128)
button_deleteFile.place(x=150+25, y=128)
button_alwaysOnTop.place(x=335,y=125)

listbox_shortcutHolder.place(x=32,y=5)

#binds and functions --
main.bind('<Button-1>', SaveLastClickPos) +main.bind('<B1-Motion>', Dragging)
label_creditCreator.bind('<Button-1>', SaveLastClickPos) +label_creditCreator.bind('<B1-Motion>', Dragging)
listbox_shortcutHolder.bind('<Button-1>', SaveLastClickPos) +listbox_shortcutHolder.bind('<B1-Motion>', Dragging)
listbox_closeApplication.bind('<Button-1>', SaveLastClickPos) +listbox_closeApplication.bind('<B1-Motion>', Dragging)

button_importFile.config(anchor="w")
button_executeFile.config(anchor="w")
button_deleteFile.config(anchor="w")

#________________________| end events |________________________#

Thread(target=detailedShavow).start()
Thread(target=autoUpdating).start()
Thread(target=autoUpdatingCLOSE).start()
Thread(target=checkOnTop).start()

insertAllPrograms()

refreshFolder()

command_recovery()

window.mainloop()

#███████╗██╗  ██╗ █████╗ ██╗   ██╗ ██████╗ ██╗    ██╗
#██╔════╝██║  ██║██╔══██╗██║   ██║██╔═══██╗██║    ██║
#███████╗███████║███████║██║   ██║██║   ██║██║ █╗ ██║
#╚════██║██╔══██║██╔══██║╚██╗ ██╔╝██║   ██║██║███╗██║
#███████║██║  ██║██║  ██║ ╚████╔╝ ╚██████╔╝╚███╔███╔╝
#╚══════╝╚═╝  ╚═╝╚═╝  ╚═╝  ╚═══╝   ╚═════╝  ╚══╝╚══╝ 
