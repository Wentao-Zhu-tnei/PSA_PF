import tkinter as tk
from tkinter import *
from tkinter import ttk

# Module 2: opens a new window with a progress bar and calls fucntion from module 3 to display info to user

def updateProgressBar(pct, text):
    global btn1

    # Set progress bar value to i / this will change the progress bar visually.
    progress["value"] = pct

    # Update label
    label2.config(text=text)

    # Update the progress bar.
    progress.update()  

    # Condition: if "i" which is the progress bar value gets to 100 then the progress bar will be full, at this point the process will be done.
    # When condition is met change button state to normal.
    # This condition is in place so the user cannot exit the process untill it is completed.
    if pct >= 100:
        btn1.config(state="normal")
        btn1["text"] = "Close"

# quit_window: Function to bypass closing the second window.
def quit_window():
    pass

def delProgressBar():
    global win2
    win2.destroy()

def createProgressBar(strTitle, bBtn):
    global win2
    global btn1
    global progress 

    win2 = tk.Tk()
    win2.title(strTitle)
    win2.attributes("-topmost", True)
    win2.resizable(width=False, height=False)
    win2.attributes("-toolwindow", 1)
    # Height and Width 
    H = int(350)
    W = int(150)
    ws = W
    hs = H
    sw = win2.winfo_screenwidth()
    sh = win2.winfo_screenheight()
    x = (sw/2) - (ws/2)
    y = (sh/2) - (hs/2)
    wstr = str(hs) + "x" + str(ws) + "+" + str(int(x)) + "+" + str(int(y))
    win2.geometry(wstr)    
    win2.wm_protocol("WM_DELETE_WINDOW", lambda: quit_window())

    # Label 2: for displaying process steps to user
    global label2
    label2 = ttk.Label(win2)
    label2.pack(pady=10)

    W1 = int(W * 2)

    # progress: Progress Bar for displaying amount of process completed to the user.
    progress = ttk.Progressbar(win2, orient=HORIZONTAL, length=W1, mode="determinate")
    progress.pack(pady=10)

    #if bBtn:
    #    strBtn = "Close"
    #else:
    #    strBtn = "Not in use"
    strBtn = "Not in use"
    btn1 = tk.Button(win2, text=strBtn, state="disabled", command=win2.destroy)
    btn1.pack(pady=10)
