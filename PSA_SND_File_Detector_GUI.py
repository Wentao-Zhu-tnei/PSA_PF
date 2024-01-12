import tkinter as tk
from tkinter import LabelFrame, Button, Text
from tkinter.messagebox import showinfo
from tkinter.messagebox import askyesno
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog
import os
import shutil
import time
from datetime import datetime
import tkinter.scrolledtext as st

# import PSA and SnD tool specific functions
import PSA_SND_Interface_Functions
import PSA_SND_Utilities as ut
import PSA_SND_FD_Constants as fd
import PSA_SND_Constants as const

PSA_dict = ut.PSA_SND_read_config("PSA_SND_config.txt")
initDir=PSA_dict["SND_PSA_FILE_DETECTOR_FOLDER"]
target_folder = initDir
dest_folder = PSA_dict["PSA_SND_ROOT_WORKING_FOLDER"] + const.strDataResults
tempPFfolder = PSA_dict["PSA_SND_ROOT_WORKING_FOLDER"] + "\\tempPF\\"

##################
#### GEOMETRY ####
##################
"""
Using three label frames for "user data entry", "control buttons" & "message log"
"""
global bRunProcess
global log_file

bRunProcess = False
log_file=""

# Default padding for spacing for data entry field
padX = 10
padY = 5

# Window #
root = tk.Tk()
root.title("PSA and S&D Tool - File Detector and Monitoring System")
w=650
h=600
ws=root.winfo_screenwidth()
hs=root.winfo_screenheight()
x=(ws/2)-(w/2)
y=(hs/2)-(h/2)
y=0
root.geometry("%dx%d+%d+%d" % (w,h,x,y))

frame1 = tk.LabelFrame(root, text="User Inputs")
frame1.grid(row=0, column=0, sticky=tk.W, padx=padX, pady=padY)

frame2 = tk.LabelFrame(root, text="Controls")
frame2.grid(row=0, column=1, sticky=tk.W, padx=padX, pady=padY)

frame3 = tk.LabelFrame(root, text="Log Messages")
frame3.grid(row=2, column=0, sticky=tk.W, padx=padX, pady=padY)

# Call backs and time repeating functions

def update_scrollTxt():
    global scrollTxt
    global bRunProcess
    global sleep_timems

    if bRunProcess:
        detect_files()
        scrollTxt.after(sleep_timems, update_scrollTxt)

def update_heartbeat():
    global scrollTxt
    global bRunProcess
    global heartbeatms
    global file_counter_PSA
    global file_counter_SND
    global file_counter_invalid
    global log_file

    if bRunProcess:
        string0 = "HEARTBEAT"
        string1 = "PSA files in the folder: " + str(file_counter_PSA)
        string2 = "SND files in the folder: " + str(file_counter_SND)
        string3 = "Invalid files in the folder: " + str(file_counter_invalid)

        log_msg(log_file,string0)
        log_msg(log_file,string1)
        log_msg(log_file, string2)
        log_msg(log_file, string3)

        scrollTxt.after(heartbeatms, update_heartbeat)


# Scroll text window for log msg #
scrollTxt = st.ScrolledText(frame3, width=60, height=25)
scrollTxt.grid(row=0, column=0, padx=padX, pady=padY, sticky=tk.W, columnspan=2)

# Annotations #
labelTargetFolderDefault = tk.Label(frame1, text=initDir).grid(row=0,column=1,sticky=tk.W,padx=padX, pady=padY)
labelSleepTime = tk.Label(frame1, text='Sleep Time (secs)').grid(row=1,column=0,sticky=tk.W,padx=padX, pady=padY)
labelHeartbeat = tk.Label(frame1, text='Heartbeat Time (mins)').grid(row=2,column=0,sticky=tk.W,padx=padX, pady=padY)
labelSubProc = tk.Label(frame1, text='Create Subprocess').grid(row=1,column=2,sticky=tk.W,padx=padX, pady=padY)

# Sleep Time #
displaySleepTime = tk.Label(frame1,text="", fg='black')
displaySleepTime.grid(row=1, column=2)

# Heartbeat #
displayHeartbeat = tk.Label(frame1,text="", fg='black')
displayHeartbeat.grid(row=2, column=2)


##################
### CALL BACKS ###
##################

# Target Folder #
def click_TargetFolderBtn():
    global target_folder
    global initDir

    target_folder = filedialog.askdirectory(initialdir=initDir)
    if target_folder == "":
        pass
    else:
        txt = target_folder
        displayTargetFolder = tk.Label(frame1, text=txt)
        displayTargetFolder.grid(row=0,column=1,sticky=tk.W,padx=padX, pady=padY)
    return target_folder

# Invalid File Button #
def click_InvalidFileBtn():
    global bInvalidFiles

    if bInvalidFiles:
        buttonOn.config(text="No")
        bInvalidFiles = False
    else:
        buttonOn.config(text="Yes")
        bInvalidFiles = True

# Start Button #
def click_StartBtn():
    # TODO: Yes or No Prompt -> ERROR msg when missing input
    global bRunProcess
    global scrollTxt
    global sleep_timems
    global heartbeatms
    global target_folder
    global log_file
    global bSubProc

    sleep_timems = dropdownSleepTimeOpt.get() * 1000  # convert user inputs from secs to msecs
    heartbeatms = int(float(dropdownHeartbeatOpt.get()) * 1000 * 60)  # convert user inputs from minutes to msecs
    if dropdownSubProcOpt.get() == "Yes":
        bSubProc = True
    else:
        bSubProc = False

    if not bRunProcess:
        if 'target_folder' not in globals() or target_folder == "":
            tk.messagebox.showinfo("Data Entry", "Please select a target folder")
        else:
            scrollTxt.delete('1.0',tk.END)
            init()
            bRunProcess=True
            msg = "START BUTTON"
            log_msg(log_file,msg)
            update_scrollTxt()
            update_heartbeat()

# Stop Button #
def click_StopBtn():
    global bRunProcess
    global scrollTxt

    if not bRunProcess:
        tk.messagebox.showinfo("Information", "No process is currently running")

    elif askyesno(title='WARNING', message='Are you sure you want to stop the process?'):
        msg = "STOP BUTTON"
        log_msg(log_file,msg)
        bRunProcess = False

# Exit Button #
def click_ExitBtn():
    if askyesno(title='WARNING', message='Are you sure you want to stop the process and kill the window?'):
        msg = "EXIT BUTTON"
        if log_file != "":
            log_msg(log_file, msg)
        root.destroy()


##################
#### BUTTONS #####
##################

# Target Folder #
buttonTargetfolder = tk.Button(frame1, text="Select target folder", command=click_TargetFolderBtn)
buttonTargetfolder.grid(row=0, column=0, sticky=tk.W, padx=padX, pady=padY)

# Sleep Time #
optionsSleepTime = [1, 10, 60]  # Options to choose length of sleep time in secs
dropdownSleepTimeOpt = tk.IntVar(frame1)
dropdownSleepTimeOpt.set(optionsSleepTime[0])
dropdownSleepTime = tk.OptionMenu(frame1, dropdownSleepTimeOpt, *optionsSleepTime)
dropdownSleepTime.grid(row=1,column=1, sticky=tk.W, padx=padX, pady=padY)

# Heartbeat Time #
optionsHeartbeat = [0.1, 5, 15]  # Options to choose gap in between heartbeats in minutes
dropdownHeartbeatOpt = tk.DoubleVar(frame1)
dropdownHeartbeatOpt.set(optionsHeartbeat[2])
dropdownHeartbeat = tk.OptionMenu(frame1, dropdownHeartbeatOpt, *optionsHeartbeat)
dropdownHeartbeat.grid(row=2,column=1, sticky=tk.W, padx=padX, pady=padY)

# Create sub-process #
optionsSubProc = ["Yes", "No"]  # Options to choose whether subprocess is created or not
dropdownSubProcOpt = tk.StringVar(frame1)
dropdownSubProcOpt.set(optionsSubProc[0])
dropdownSubProc = tk.OptionMenu(frame1, dropdownSubProcOpt, *optionsSubProc)
dropdownSubProc.grid(row=1,column=3, sticky=tk.W, padx=padX, pady=padY)


# Start Button #
buttonStart = tk.Button(frame2, text="START", command=click_StartBtn, bg='green', fg='white',height=1, width=5)
buttonStart.grid(row=0,column=0, padx=padX, pady=padY)

# Stop Button #
buttonStop = tk.Button(frame2, text="STOP", command=click_StopBtn, bg='red', fg='white',height=1, width=5)
buttonStop.grid(row=1,column=0,padx=padX, pady=padY)

# Exit Button #
buttonExit = tk.Button(frame2, text="EXIT", command=click_ExitBtn, bg='black', fg='white',height=1, width=5)
buttonExit.grid(row=2,column=0, padx=padX, pady=padY)

# Invalid File Button #
global bInvalidFiles
bInvalidFiles = True
labelDisplay = tk.Label(frame1, text='Display invalid files? Yes/No').grid(row=2,column=2,sticky=tk.W,padx=padX, pady=padY)

buttonOn = tk.Button(frame1, text="Yes", command=click_InvalidFileBtn)
buttonOn.grid(row=2,column=3, sticky=tk.W, padx=padX, pady=padY)

##################
#### ENGINE #####
##################

def log_msg(log_fle, msg):
    """
    Create log messages in "log.txt"
    """
    global strTxt
    global scrollTxt

    now = datetime.now()
    d1 = now.strftime("%Y-%m-%d %H:%M:%S")
    f = open(log_fle, 'a')
    strTxt = d1 + " " + msg + "\n"
    f.write(strTxt)
    scrollTxt.insert(tk.END, strTxt)
    scrollTxt.see("end")

def init():
    global stop_file
    global log_file
    global target_folder
    global bRunProcess

    # Create log file
    stop_file = target_folder + "\\stop.txt"
    log_dir = '\\Log'
    log_dir = target_folder + log_dir
    log_file_name = time.strftime("%Y-%m-%d-%H-%M-%S") + ".txt"
    log_file = log_dir + "\\" + log_file_name

    bRunProcess = True

    if os.path.isfile(os.path.join(target_folder, stop_file)):
        bRunProcess = False
        msg = "Stop file detected!"
        log_msg(log_file, msg)

    if os.path.isdir(log_dir):
        pass

    else:
        os.mkdir(log_dir)
        msg = "Log folder does not exist; a new one has been created"
        log_msg(log_file, msg)


def purgePFworkspaces():
    for item in os.listdir(tempPFfolder):
        deldir = tempPFfolder + "\\" + item
        if os.path.isdir(deldir):
            if "workspace" in deldir:
                try:
                    shutil.rmtree(deldir)
                except:
                    pass


def detect_files():
    """
    Run the file detection function within given target folder
    """
    global target_folder
    global stop_file
    global log_file
    global file_counter_PSA
    global file_counter_SND
    global file_counter_invalid
    global bInvalidFiles
    global bSubProc


    file_counter_PSA = 0
    file_counter_SND = 0
    file_counter_invalid = 0

    for item in os.listdir(target_folder):
        if (fd.strPSANoFlexReqts in item) or \
            (fd.strPSAFlexReqts in item) or \
            (fd.strPSASFResponses in item) or \
            (fd.strPSASFContracts in item) or \
            (fd.strPSACandidateResponses in item) or \
            (fd.strPSACandidateContracts in item):
            file_counter_PSA += 1
            msg = item + " detected"
            PSA_SND_Interface_Functions.SND_PROCESS(target_folder, dest_folder, item, bSubProc)

        elif (fd.strSNDResponses in item) or \
            (fd.strSNDContracts) in item or \
            (fd.strSNDCandidateResponses in item) or \
            (fd.strSNDCandidateContracts in item):
            file_counter_SND += 1
            msg = item + " detected"
            PSA_SND_Interface_Functions.PSA_PROCESS(target_folder, dest_folder, item, bSubProc)
            purgePFworkspaces()

        elif os.path.isfile(os.path.join(target_folder, item)):
            file_counter_invalid += 1
            if bInvalidFiles:
                PSA_SND_Interface_Functions.move_file(target_folder, dest_folder, item, False)
                msg = "Unknown files detected! " + "[" + str(item) + "]"
            else:
                msg = ""

        else:
            msg = ""

        if msg != "":
            log_msg(log_file, msg)


##################
####WINDOW OPT####
##################

root.protocol("WM_DELETE_WINDOW", click_ExitBtn)

root.mainloop()
