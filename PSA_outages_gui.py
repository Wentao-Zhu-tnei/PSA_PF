#import tkinter as tk
import PSA_SND_Constants as const
import PSA_SND_Utilities as ut
from tkinter import *
from tkinter import filedialog as fd
from tkinter import ttk
# from conda prompt
# pip install tkcalendar 
from tkcalendar import *
from datetime import datetime, timedelta

# root.resizable(False, False)
# root.attributes('-topmost', 1)
# root.iconbitmap('<file extension>')



self = tk.Tk()
ct = datetime.now() # current_time
f = ('TkDefaultFont', 12)
i = 0

self.title('PSA Maint Generator')
#self.geometry('600x600+50+50')

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Hour Entry Widget'''
class HourEntry(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.hourstr=tk.StringVar(self,'10')
        self.hour = tk.Spinbox(self,from_=0,to=23,wrap=True,textvariable=self.hourstr,width=2,font=f,state="readonly")
        self.minstr=tk.StringVar(self,'30')
        self.min = tk.Spinbox(self,from_=0,to=59,wrap=True,textvariable=self.minstr,width=2,font=f,state="readonly")

        self.hour.grid(column=0, row=0)
        self.min.grid(column=1, row=0)

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''While shift is pressed, Minute and Hour increment increases for convenience sake'''
running = True
def start_motor(event):
    global running
    if running:
        te1.hour.config(increment=6)
        te1.min.config(increment=15)

        te2.hour.config(increment=6)
        te2.min.config(increment=15)
        
        running = False

def stop_motor(event):
    global running  
    te1.hour.config(increment=1)
    te1.min.config(increment=1)
    
    te2.hour.config(increment=1)
    te2.min.config(increment=1)

    running = True

self.bind('<Shift_L>', start_motor)
self.bind('<KeyRelease-Shift_L>', stop_motor)
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Functions for retrieving and enacting on input'''

'''
def onbtnclicked(): 
    # Initialdir = which folder you want the window to open at, title = window title, filetypes = files you want to choose 
    fileLoc = filedialog.askopenfilename(initialdir="C:\\", title="Open File", filetypes=[('Excel Files', '.xlsx')])
    dir_box['state'] = 'normal'
    dir_box.delete('1.0', 'end')
    dir_box.insert('1.0', fileLoc)
    dir_box['state'] = 'disabled'
'''

def fileBtn_clicked_TB(self, fLabel, fpat, fBtn, bInput):
    ftypes = [('Excel files', '.xlsx')]
    
    if bInput:
        file_loc = self.dict_config[const.strWorkingFolder] + '/data/input'
    else:
        file_loc = self.dict_config[const.strWorkingFolder] + '/data/results'
    
    filename = fd.askopenfilename(title='Select file', initialdir=file_loc, filetypes=ftypes, initialfile=fpat)
    fnme_extnd = filename.replace(self.dict_config[const.strWorkingFolder], '')

    if filename != '':
        fLabel['state'] = 'normal'
        fLabel.delete("1.0", "end")
        fLabel.insert("1.0", fnme_extnd)
        fLabel['state'] = 'disabled'
    if bool(fBtn):
        if (fLabel.get("1.0", "end").strip() != const.strNoFileDisp) and (
                fLabel.get("1.0", "end").strip() != const.strNoFileDisp):
            fBtn['state'] = 'normal'
        else:
            fBtn['state'] = 'disabled'

def add_zero(value):
    return value if int(value) > 9 else f'0{value}'

def open_File(path):
    if os.path.isfile(path):
        os.startfile(path)
    else:
        print('Unable to Open File')

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Functions for input processing'''
def add_outage():
    global i
    start_date = datetime.strftime(de1.get_date(), '%Y-%m-%d') + f' {add_zero(te1.hour.get())}:{add_zero(te1.min.get())}:00'
    end_date = datetime.strftime(de2.get_date(), '%Y-%m-%d') + f' {add_zero(te2.hour.get())}:{add_zero(te2.min.get())}:00'

    days_difference = (datetime.strptime(end_date, '%Y-%m-%d %H:%M:%S') - datetime.strptime(start_date, '%Y-%m-%d %H:%M:%S')).days
    minutes_difference = (datetime.strptime(end_date[-8:], '%H:%M:%S') - datetime.strptime(start_date[-8:], '%H:%M:%S')).seconds / 86400
    difference = days_difference + minutes_difference
    if difference <= 0:
        label0['text'] = 'Enter a Future Date'
        return 

    asset_id = value.get()
    if asset_id == '':
        label0['text'] = 'Select an Asset'
        return 

    asset_type = const.strAssetLine if asset_id[-1].isnumeric() else const.strAssetTrans
    description = description_entry.get('1.0','end')[:-1]

    file = dir_box.get('1.0', 'end')[:-1]

    if '.xlsx' not in file:
        df = pd.DataFrame(columns=['ASSET_LONG_NAME','ASSET_ID','ASSET_TYPE','START_OUTAGE','END_OUTAGE','DURATION','DESCRIPTION'])
        file += f'{const.strMaintFile}.xlsx'
        if not messagebox.askokcancel('Info', f'This will generate a new file in:\n{file}'):
            label0['text'] = 'Addition Cancelled'
            return
    else:
        df = pd.read_excel(file)
        if df.shape[1] != 7:
            label0['text'] = 'File Shape Error, aborting'
            return
    
    df.loc[len(df)] = [asset_id, asset_id, asset_type, start_date, end_date, difference, description]
    try:
        pa.auto_expand_and_save(df, file, index=False, sheet_name=const.strDataSheet)
        #df.to_excel(file, sheet_name=const.strDataSheet, index=False)
        print(f'Saved to: {file}')
        
    except:
        label0['text'] = 'File Save Error'

    i += 1
    label0['text'] = 'Outage Added'
    return 

def undo():
    global i
    if i > 0:
        file = dir_box.get('1.0', 'end')[:-1]

        df = pd.read_excel(file)      
        if df.shape[1] != 7:
            label0['text'] = 'File Shape Error'
            return
        
        df.drop(index=len(df)-1)
        pa.auto_expand_and_save(df, file, index=False, sheet_name=const.strDataSheet)
        label0['text'] = 'Addition Undone'

        i -= 1

    else:
        label0['text'] = 'Nothing to Undo'

    return 

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Data Entry Parameters'''

# -----------------------------------------------------------
'''Frame 0: Message Box'''
frame0 = ttk.Frame(self, borderwidth=5, relief='groove')
label0 = tk.Label(frame0, text='', font=f)
label0.pack(fill=tk.X)

# -----------------------------------------------------------
'''Frame 1: Select Maint File, Select Asset, Enter Description'''
self.frame1 = ttk.Frame(self)
self.label1 = tk.Label(self.frame1, text='Select Maintenance File')
self.label1.grid(column=1, row=0, pady=5)
self.button1 = tk.Button(self.frame1, text="Browse", command=lambda: fileBtn_clicked_TB(self, self.dir_box, ' ', self.button3, True))
self.button1.grid(column=0, row=1, padx=10, pady=5)

dict_config = ut.PSA_SND_read_config(const.strConfigFile)
pic_excel_logo = tk.PhotoImage(file=dict_config[const.strWorkingFolder] + "/tkinterapp/assets/Excel_Logo.png").subsample(120)
self.button3 = ttk.Button(self.frame1, text="Edit Data", command=lambda: open_File(self.dir_box.get('1.0','end')[:-1]), image=pic_excel_logo, compound=tk.LEFT, state=tk.NORMAL)
self.button3.grid(column=2, row=1, padx=10, pady=5)

self.dir_box = tk.Text(self.frame1, width=66, height=1)
self.dir_box['state'] = 'normal'
self.dir_box.insert('1.0', const.strNoFileDisp)
self.dir_box['state'] = 'disabled'
self.dir_box.grid(column=1, row=1, pady=5)

Asset_label = tk.Label(self.frame1, text='Select Asset')
Asset_label.grid(column=1, row=2, padx=10, pady=10, sticky=tk.W)
Assets_list = (0,1,2,3,4,5,6)
value = tk.StringVar()
ae1 = ttk.Combobox(self.frame1, values=Assets_list, textvariable=value)
ae1.grid(column=0, row=3, columnspan=3, padx=10, pady=0, sticky=tk.E)

tk.Label(self.frame1, text='Description').grid(column=1, row=4, pady=5)
description_entry = tk.Text(self.frame1, height=3)
description_entry.insert('1.0', 'Planned outage for maintenance')
description_entry.grid(column=0, row=5, columnspan=3, pady=5)

# -----------------------------------------------------------
'''Frame 3: Initial Start Time/Date Entry'''
frame3 = ttk.Frame(self)
tk.Label(frame3, text='Start Time:').pack(padx=10, expand=True)
de1 = DateEntry(frame3, year=ct.year, month=ct.month, day=ct.day,
                 date_pattern='dd/mm/yy',
                 selectbackground='gray80',
                 selectforeground='black',
                 normalbackground='white',
                 normalforeground='black',
                 background='gray90',
                 foreground='black',
                 bordercolor='gray90',
                 othermonthforeground='gray50',
                 othermonthbackground='white',
                 othermonthweforeground='gray50',
                 othermonthwebackground='white',
                 weekendbackground='white',
                 weekendforeground='black',
                 headersbackground='white',
                 headersforeground='gray70')
de1.pack()

tk.Label(frame3, text='Hour | Min').pack(padx=10)
te1 = HourEntry(frame3)
te1.pack()


frame2 = ttk.Frame(self)
tk.Label(frame2, text='End Time:').pack(padx=10)
de2 = DateEntry(frame2, year=ct.year, month=ct.month, day=ct.day,
                 date_pattern='dd/mm/yy',
                 selectbackground='gray80',
                 selectforeground='black',
                 normalbackground='white',
                 normalforeground='black',
                 background='gray90',
                 foreground='black',
                 bordercolor='gray90',
                 othermonthforeground='gray50',
                 othermonthbackground='white',
                 othermonthweforeground='gray50',
                 othermonthwebackground='white',
                 weekendbackground='white', 
                 weekendforeground='black',
                 headersbackground='white',
                 headersforeground='gray70')     
de2.pack()

tk.Label(frame2, text='Hour | Min').pack(padx=10)
te2 = HourEntry(frame2)
te2.pack()

#frame0.grid(column=0, row=0, columnspan=2)
self.frame1.grid(column=0, row=1, columnspan=2)
frame3.grid(column=0, row=2, pady=15)
frame2.grid(column=1, row=2, pady=15)

tk.Button(text='Add Outage', command=add_outage, bg='green', fg='white').grid(pady=10, column=0, row=3, sticky=tk.E) 
tk.Button(text='Undo', command=undo, bg='black', fg='white').grid(pady=30, column=1, row=3, sticky=tk.W)

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Loop'''

self.mainloop()