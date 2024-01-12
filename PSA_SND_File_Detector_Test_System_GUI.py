import tkinter as tk
from tkinter import LabelFrame, Button, Text
from tkinter.messagebox import showinfo
from tkinter.messagebox import askyesno
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog as fd
import time
import pandas as pd
import PSA_SND_Utilities as ut
import PSA_SND_FD_Constants as fdc

root=tk.Tk()
#centre window on screen
w=730
h=580
ws=root.winfo_screenwidth()
hs=root.winfo_screenheight()
x=(ws/2)-(w/2)
y=(hs/2)-(h/2)
root.geometry("%dx%d+%d+%d" % (w,h,x,y))
root.title('PSA and S&D Tool - File Detector System (Test File Creation)')

finished=True
PSA_dict = ut.PSA_SND_read_config("PSA_SND_config.txt")
initDir=PSA_dict["PSA_SND_FILE_DETECTOR_FOLDER"]
tgtFolder="No folder selected"

def write_file(fstr, tstr):
    global st
    global tgtFolder
    fname = tgtFolder + "\\" + fstr
    testData = [fname]
    df=pd.DataFrame(testData)
    with pd.ExcelWriter(fname) as writer:  
         df.to_excel(writer, sheet_name = 'Test Data',header = True, index= True)
    st.insert(tk.END,tstr + " " + fstr + "\n")
    st.see("end")
    
def create_files():
    tstr = time.strftime("%Y-%m-%d-%H-%M-%S")
    if tstr[-1:]=="0" or \
        tstr[-1:]=="2" or \
        tstr[-1:]=="4" or \
        tstr[-1:]=="6" or \
        tstr[-1:]=="8":
            rstr = "AUT_" + tstr
    else:
        rstr = "MAN_" + tstr

    if tstr[-1:] == "0":  
        fstr = rstr + "_" + fdc.strPSANoFlexReqts + ".xlsx"    
    elif tstr[-1:]=="1":
        fstr = rstr + "_" + fdc.strPSAFlexReqts + ".xlsx"        
    elif tstr[-1:]=="2":
        fstr = rstr + "_" + fdc.strSNDResponses + ".xlsx"        
    elif tstr[-1:]=="3":
        fstr = rstr + "_" + fdc.strPSASFResponses + ".xlsx"        
    elif tstr[-1:]=="4":
        fstr = rstr + "_" + fdc.strSNDCandidateResponses + "-Vn.xlsx"
    elif tstr[-1:]=="5":
        fstr = rstr + "_" + fdc.strPSACandidateResponses + "-Vn.xlsx"
    elif tstr[-1:]=="6":
        fstr = rstr + "_" + fdc.strSNDContracts + ".xlsx"        
    elif tstr[-1:]=="7":
        fstr = rstr + "_" + fdc.strPSASFContracts + ".xlsx"        
    elif tstr[-1:]=="8":
        fstr = rstr + "_" + fdc.strSNDCandidateContracts + "-Vn.xlsx"
    elif tstr[-1:]=="9":
        fstr = rstr + "_" + fdc.strPSACandidateContracts + "-Vn.xlsx"
    print(str(rstr))
    
    write_file(fstr, tstr)
    

def click_StartBtn():
    global finished
    global st
    global sleepTime

    if tgtFolder != "No folder selected":
        sleepTime = dropdownSleepTimeOpt.get() * 1000  # convert user inputs from secs to msecs
        if finished:
            finished=False
            st.insert(tk.END,"START BUTTON\n")
            st.see("end")
            st_update()
    else:
        showinfo("Data Entry","Please select a target folder")
    
def click_StopBtn():
    global finished
    global st
    st.insert(tk.END,"STOP BUTTON\n")
    st.see("end")
    finished=True

def click_ExitBtn():
    if askyesno(title='WARNING', message='Are you sure you want to stop the process and kill the window?'):
        root.destroy()
    
def st_update():
    global finished
    global st
    global sleepTime
    if not finished:
        create_files()
        st.after(sleepTime,st_update)


def click_folderBtn():
    global fname
    global tgtFolder
    str=fd.askdirectory(title='Select target folder', initialdir=initDir)
    if str != "": 
        tgtFolder = str
        fname['state']='normal' 
        fname.delete("1.0","end")
        fname.insert("1.0",tgtFolder)
        fname['state']='disabled'

lf1=LabelFrame(root, text="Control Panel")
lf1.grid(column=0, row=0, padx=10, pady=10, sticky=tk.NW, columnspan=2)
folderBtn=Button(lf1, text="Select Target Folder", command=click_folderBtn)
fname=Text(lf1, width=65, height=1)
fname['state']='normal' 
fname.insert("1.0",tgtFolder)
fname['state']='disabled' 
folderBtn.grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
fname.grid(row=0, column=1, padx=10, pady=10, sticky=tk.W)

# Sleep Time #
optionsSleepTime = [6, 60, 600]  # Options to choose length of sleep time in secs
dropdownSleepTimeOpt = tk.IntVar(lf1)
dropdownSleepTimeOpt.set(optionsSleepTime[0])
dropdownSleepTime = tk.OptionMenu(lf1, dropdownSleepTimeOpt, *optionsSleepTime)
dropdownSleepTime.grid(row=1,column=1, sticky=tk.W, padx=10, pady=10)
labelSleepTime = tk.Label(lf1, text='Seep time (secs)').grid(row=1,column=0,sticky=tk.W,padx=10, pady=10)

lf2=LabelFrame(root, text="Control Buttons")
lf2.grid(column=0, row=1, padx=10, pady=10, sticky=tk.NW)

btn1=Button(lf2, text="START", command=click_StartBtn, bg='green', fg='white',height=1, width=5)
btn1.grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)

btn2=Button(lf2, text="STOP", command=click_StopBtn, bg='red', fg='white',height=1, width=5)
btn2.grid(row=1, column=1, padx=10, pady=10, sticky=tk.W)

btn3=Button(lf2, text="EXIT", command=click_ExitBtn, bg='black', fg='white',height=1, width=5)
btn3.grid(row=1, column=2, padx=10, pady=10, sticky=tk.W)

lf3=LabelFrame(root, text="Files Created")
lf3.grid(column=0, row=2, padx=10, pady=10, sticky=tk.NW, columnspan=2)
st=ScrolledText(lf3, width=80, height=20)
st.grid(column=0, row=0, padx=10, pady=10)
st.after(1000,st_update)

root.protocol("WM_DELETE_WINDOW", click_ExitBtn)
root.mainloop()





