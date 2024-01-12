import tkinter as tk
import pandas as pd
import PSA_SND_Constants as const
import PSA_SND_Utilities as ut
import PSA_analysis as pa
import os

from tkinter import ttk, messagebox, Widget
from tkinter import filedialog as fd
from tkcalendar import DateEntry
from datetime import datetime, timedelta
from PIL import Image, ImageTk

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Hour Entry Widget'''
class HourEntry(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        f = ('TkDefaultFont', 12)
        self.hourstr=tk.StringVar(self,'12')
        self.hour = tk.Spinbox(self,from_=0,to=23,wrap=True,textvariable=self.hourstr,width=2,font=f,state="readonly")
        self.minstr=tk.StringVar(self,'00')
        self.min = tk.Spinbox(self,from_=0,to=59,wrap=True,textvariable=self.minstr,width=2,font=f,state="readonly")

        self.hour.grid(column=0, row=0)
        self.min.grid(column=1, row=0)

class ExcelEntry(tk.Frame):
    def __init__(self, parent, checkbox1=None, checkbox2=None, description=None):
        super().__init__(parent)
        self.parent = parent
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=4)
        self.grid_columnconfigure(2, weight=1)

        self.grid_rowconfigure(0, weight=1)

        self.dict_config = ut.PSA_SND_read_config(const.strConfigFile)
        self.pic_excel_logo = tk.PhotoImage(file=self.dict_config[const.strWorkingFolder] + "/tkinterapp/assets/Excel_Logo.png").subsample(120)

        self.browse1_btn = tk.Button(self, text="Browse", command=lambda: self.fileBtn_clicked_TB(self.dir_box, '', self.editdata1_btn, True, checkbox1, checkbox2, description))
        self.browse1_btn.grid(column=0, row=0, padx=10, ipadx=20)

        self.dir_box = tk.Text(self, width=60, height=1)
        self.dir_box.insert('1.0', const.strNoFileDisp)
        self.dir_box['state'] = 'disabled'
        self.dir_box.grid(column=1, row=0, padx=10)

        self.editdata1_btn = ttk.Button(self, text="Edit Data", command=lambda: ut.open_xl(self.dir_box, self.dict_config[const.strWorkingFolder]), image=self.pic_excel_logo, compound=tk.LEFT, state=tk.DISABLED)
        self.editdata1_btn.grid(column=2, row=0, padx=10)

    def fileBtn_clicked_TB(self, fLabel, fpat, fBtn, bInput, checkbox1, checkbox2, description):
        ftypes = [('Excel files', '.xlsx')]
        fLabel['state'] = 'normal'
        fLabel.delete("1.0", "end")
        fLabel.insert("1.0", const.strNoFileDisp)
        fLabel['state'] = 'disabled'
        
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
            if bool(checkbox1) and bool(checkbox2) and bool(description):
                description.delete('1.0','end')
                if 'MAINT' in fnme_extnd.upper():
                    checkbox1.select()
                    checkbox2.deselect()
                    description.insert('1.0', 'Planned Outage for Maintenance')
                else: 
                    checkbox1.deselect()
                    checkbox2.select()
                    description.insert('1.0', 'Contingency outage')
                checkbox1['state'] = 'disabled'
                checkbox2['state'] = 'disabled'

        if bool(fBtn):
            if (fLabel.get("1.0", "end").strip() != const.strNoFileDisp):
                fBtn['state'] = 'normal'
            else:
                fBtn['state'] = 'disabled'

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Main App'''
class App(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        # -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '''Data Entry Parameters'''
        self.title('PSA Maint Generator')
        self.protocol("WM_DELETE_WINDOW", self.quit_me) 
        # Because of an issue with matplotlib, the program would actually keep running after the exit button is closed while a graph is displayed
        # To counter this, we set protocol to both quit and destroy when the exit button is pressed. 

        width = self.winfo_screenwidth() 
        height = self.winfo_screenheight()
        #self.state('normal') #normal, iconic, withdrawn, or zoomed
        self.geometry(f'1520x640+{(width-1500)//2}+{(height-620)//2}')
        
        self.dict_config = ut.PSA_SND_read_config(const.strConfigFile)
        asset_value = tk.StringVar()

        self.ct = datetime.now()
        
        self.asset_dict = [] # Used for storing asset_types
        self.running = True
        self.df_backup = [] # Used for storing one df back in time
        self.i = 0 # Used for add outage undo function
        self.v = 0 # Used for update outages undo function

        self.bind('<Shift_L>', self.start_motor)
        self.bind('<KeyRelease-Shift_L>', self.stop_motor)
        #self.bind('<KeyPress-F5>', self.refresh_image) # Doesn't work because two arguments given when only one is taken (self, event) vs (self)

        self.LabelFrame00 = ttk.LabelFrame(self, text='Add New Outage') # Data Entry
        frame6 = tk.Frame(self)
        self.LabelFrame10 = ttk.LabelFrame(frame6, text='Resulting Graph') # Plot Show
        self.LabelFrame20 = ttk.LabelFrame(frame6, text='Update Existing Outages')
        # -----------------------------------------------------------
        '''Frame 01: Select Asset File'''
        self.frame01 = ttk.LabelFrame(self.LabelFrame00, text='Select Asset File')

        self.fe2 = ExcelEntry(self.frame01)
        self.fe2.pack(padx=10, pady=10)
        # -----------------------------------------------------------
        '''Frame 1: Select Asset, Enter Description'''
        frame1 = ttk.Frame(self.LabelFrame00)

        self.select_asset = tk.Label(frame1, text='Select Asset')
        self.select_asset.pack(padx=10, pady=10)
        #self.select_asset.grid(column=0, row=2, columnspan=3, padx=10, ipady=10, sticky=tk.S)

        asset_list = ('No_Asset_File')
        self.pick_asset = ttk.Combobox(frame1, values=asset_list, textvariable=asset_value, width = 35, postcommand=lambda: self.pick_asset.configure(values=self.update_list()))
        self.pick_asset.pack(padx=10, pady=0)
        #self.pick_asset.grid(column=0, row=3, columnspan=3, padx=10, pady=0)

        self.description = tk.Label(frame1, text='Reason For Outage')
        self.description.pack(padx=10, pady=10)
        #self.description.grid(column=0, row=4, columnspan=3, padx=10, ipady=10, sticky=tk.S)

        self.description_entry = tk.Text(frame1, height=3)
        self.description_entry.pack(padx=10, pady=0)
        #self.description_entry.grid(column=0, row=5, columnspan=3, padx=10, pady=0)

        # -----------------------------------------------------------
        '''Frame 0: Select Maint File'''
        self.frame0 = ttk.LabelFrame(self, text='Select Outage File')
    
        self.checkbox_var1 = tk.StringVar()
        self.maint_checkbox1 = tk.Checkbutton(self.frame0, text='Ouput Maintenance File', command=lambda: self.AF1_checkbox(self.maint_checkbox1), variable=self.checkbox_var1, onvalue='MAINT', offvalue='CONT')
        self.maint_checkbox1.select()
        self.maint_checkbox1.grid(padx=10, pady=5, column=0, row=0)

        self.checkbox_var2 = tk.StringVar()
        self.maint_checkbox2 = tk.Checkbutton(self.frame0, text='Ouput Contingency File', command=lambda: self.AF1_checkbox(self.maint_checkbox2), variable=self.checkbox_var2, onvalue='MAINT', offvalue='CONT')
        self.maint_checkbox2.select()
        self.maint_checkbox2.grid(padx=10, pady=5, column=0, row=1)

        self.fe1 = ExcelEntry(self.frame0, self.maint_checkbox1, self.maint_checkbox2, self.description_entry)
        self.bind('<FocusOut>', self.clear_canvas)

        self.fe1.grid(padx=10, pady=10, column=1, row=0, rowspan=2)
        self.check_changed()
        
        # -----------------------------------------------------------
        '''Frame 2: Initial Start Time/Date Entry'''
        frame4 = ttk.LabelFrame(self.LabelFrame00, text='')
        self.UTC_label = tk.Label(frame4, text=f'Updating')
        self.UTC_label.grid(row=0, column=0, columnspan=3)
        self.update_utc(self.UTC_label)
        frame2 = ttk.Frame(frame4)

        self.start_date = tk.Label(frame2, text='Start Time:')
        self.start_date.pack(padx=10, expand=True)
        self.de1 = DateEntry(frame2, year=self.ct.year, month=self.ct.month, day=self.ct.day,
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
        self.de1.pack()

        self.start_time = tk.Label(frame2, text='Hour | Min')
        self.start_time.pack(padx=10)
        self.te1 = HourEntry(frame2)
        self.te1.pack()
        
        frame2.grid(column=0, row=1, padx=10, pady=10, sticky=tk.W)

        seperator = tk.Label(frame4, text='--->')
        seperator.grid(column=1, row=1, padx=10, pady=10)

        # -----------------------------------------------------------
        '''Frame 3: Final end Time/Date Entry'''
        frame3 = ttk.Frame(frame4)
        self.end_date = tk.Label(frame3, text='End Time:')
        self.end_date.pack(padx=10)
        self.de2 = DateEntry(frame3, year=self.ct.year, month=self.ct.month, day=self.ct.day,
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
        self.de2.pack()

        self.end_time = tk.Label(frame3, text='Hour | Min')
        self.end_time.pack(padx=10)
        self.te2 = HourEntry(frame3)
        self.te2.pack()

        frame3.grid(column=2, row=1, padx=10, pady=10, sticky=tk.W)

        # -----------------------------------------------------------
        '''Frame 5: Update Existing Outages'''
        frame5 = tk.Frame(self.LabelFrame20)

        self.update_description = tk.Label(frame5, text='Add the following timedelta to all\noutages in the Maintenance File')

        self.delta_t = tk.Label(frame5, text='Days | Hours | Mins')
        self.te3 = HourEntry(frame5)
        self.te3.hourstr.set('0')

        self.daystr = tk.StringVar(self,'7')
        self.de3 = tk.Spinbox(frame5,from_=0,to=364,wrap=True,textvariable=self.daystr,width=2,font=('TkDefaultFont', 12),state="readonly")

        self.update_outage_btn = tk.Button(frame5, text='Update Outages', command=self.update_outage, bg='green', fg='white')
        self.update_undo_btn = tk.Button(frame5, text='Undo', command=self.restore_file, bg='black', fg='white')

        self.update_description.grid(padx=30, pady=10, column=0, row=0, rowspan=2)
        self.delta_t.grid(pady=10, column=1, row=0, columnspan=2)
        self.te3.grid(column=2, row=1)
        self.de3.grid(column=1, row=1)
        self.update_outage_btn.grid(padx=30, pady=10, column=3, row=0, rowspan=2, sticky=tk.E)
        self.update_undo_btn.grid(pady=0, column=4, row=0, rowspan=2, sticky=tk.W)

        frame5.grid(column=0, row=0, padx=10, pady=10, sticky=tk.N)

        # -----------------------------------------------------------
        '''Grid Packing And Final Buttons for self.LabelFrame00'''
        self.frame0.grid(column=0, row=0, columnspan=2, padx=10, pady=10)
        self.frame01.grid(column=0, row=1, columnspan=2, padx=10, pady=10)
        frame1.grid(column=0, row=2, columnspan=2, padx=10, pady=10)
        frame4.grid(column=0, row=3, columnspan=2, padx=10, pady=10)

        self.add_outage_btn = tk.Button(self.LabelFrame00, text='Add Outage', command=self.add_outage, bg='green', fg='white')
        self.add_outage_btn.grid(pady=10, column=0, row=4, sticky=tk.E) 
        self.undo_btn = tk.Button(self.LabelFrame00, text='Undo', command=self.undo, bg='black', fg='white')
        self.undo_btn.grid(pady=30, column=1, row=4, sticky=tk.W)

        # -----------------------------------------------------------
        '''Grid Packing And Final Buttons for self.LabelFrame10'''
        self.cy = 300 # Height of Canvas area
        self.cx = 664 # Width of Canvas area

        self.Future_days = tk.Label(self.LabelFrame10, text='Timespan (Days):')
        self.Future_days.grid(pady=10, padx=5, column=0, row=0, sticky=tk.E)
        
        self.days_no = tk.IntVar(self.LabelFrame10, 10)
        self.Future_days_no = tk.Entry(self.LabelFrame10, textvariable=self.days_no)
        self.Future_days_no.grid(pady=10, padx=5, column=1, row=0, sticky=tk.W)

        self.focus_btn = tk.Button(self.LabelFrame10, text='Fullscreen', command=self.image_focus)
        self.focus_btn.grid(pady=10, padx=10, column=2, row=0, sticky=tk.W, ipadx=20)
        self.refresh_btn = tk.Button(self.LabelFrame10, text='Refresh Graph', command=self.refresh_image, bg='blue', fg='white')
        self.refresh_btn.grid(pady=10, padx=10, column=2, row=0, sticky=tk.E)
        self.plot_label = tk.Canvas(self.LabelFrame10, height=self.cy, width=self.cx, bg='white')
        self.plot_label.grid(pady=10, padx=10, column=0, columnspan=3, row=1, sticky=tk.N)

        # -----------------------------------------------------------
        '''Final Packing'''
        self.LabelFrame20.grid(padx=10, pady=0, column=1, row=1, sticky=tk.NW) # Update Outages
        self.LabelFrame10.grid(padx=10, pady=20, column=1, row=2, sticky=tk.NW) # Plot Show

        self.frame0.grid(padx=10, pady=10, column=0, row=0, columnspan=2, sticky=tk.S)
        frame6.grid(padx=10, pady=0, column=1, row=1, sticky=tk.NW)
        self.LabelFrame00.grid(padx=10, pady=0, column=0, row=1, sticky=tk.NE) # Add New Outage
        
      
        

        #self.LabelFrame00.pack(padx=10, pady=10, fill=tk.BOTH, expand=True, side=tk.LEFT) # Data Entry
        #self.LabelFrame10.pack(padx=10, pady=10, fill=tk.BOTH, expand=True, side=tk.RIGHT) # Plot Show

    # -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '''While shift is pressed, Minute and Hour increment increases for convenience sake'''
    def start_motor(self, event):
        if self.running:
            self.te1.hour.config(increment=6)
            self.te1.min.config(increment=15)

            self.te2.hour.config(increment=6)
            self.te2.min.config(increment=15)
            self.running = False

    def stop_motor(self, event): 
        self.te1.hour.config(increment=1)
        self.te1.min.config(increment=1)
        
        self.te2.hour.config(increment=1)
        self.te2.min.config(increment=1)
        self.running = True

    # -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '''Functions for input processing'''
    def AF1_checkbox(self, chkbox):
        self.maint_checkbox1.deselect()
        self.maint_checkbox2.deselect()
        chkbox.select()
        self.check_changed()

    def check_changed(self, event=None):
        if self.fe1.dir_box.get("1.0", "end").strip() == const.strNoFileDisp:
            self.description_entry.delete('1.0', 'end')
            if self.checkbox_var1.get() == 'MAINT':
                self.description_entry.insert('1.0', 'Planned Outage for Maintenance')
                self.maint_checkbox2.deselect()
            else:
                self.description_entry.insert('1.0', 'Contingency outage')
                self.maint_checkbox1.deselect()

    def image_focus(self):
        self.frame0.grid_forget()
        self.LabelFrame00.grid_forget()
        self.LabelFrame10.grid_forget()
        self.LabelFrame20.grid_forget()

        self.cx = 1096
        self.cy = 495

        self.plot_label.configure(height=self.cy, width=self.cx)
        self.focus_btn.configure(text='Exit Fullscreen', command=self.exit_fullscreen)
        self.refresh_image()
        
        self.LabelFrame10.pack(anchor=tk.E)
        return
    
    def update_utc(self, label):
        label.configure(text=f"Times are interpreted as UTC+00:00.\nUTC time now: {datetime.utcnow().strftime('%H:%M')}")
        self.after(1000, lambda: self.update_utc(self.UTC_label))
        return
    
    def exit_fullscreen(self):
        self.LabelFrame10.pack_forget()

        self.cx = 664
        self.cy = 300

        self.plot_label.configure(height=self.cy, width=self.cx)
        self.focus_btn.configure(text='Fullscreen', command=self.image_focus)
        self.refresh_image()

        self.LabelFrame20.grid(padx=10, pady=0, column=1, row=1, sticky=tk.NW) # Update Outages
        self.LabelFrame10.grid(padx=10, pady=20, column=1, row=2, sticky=tk.NW) # Plot Show
        self.frame0.grid(padx=10, pady=10, column=0, row=0, columnspan=2, sticky=tk.S) # Select Maint file
        self.LabelFrame00.grid(padx=10, pady=0, column=0, row=1, sticky=tk.NE) # Add New Outage
        return

    def clear_canvas(self, event=None):
        if self.fe1.dir_box.get("1.0", "end").strip() == const.strNoFileDisp:
            self.plot_label.delete('all')
            self.maint_checkbox1['state'] = 'normal'
            self.maint_checkbox2['state'] = 'normal'
            self.fe2.browse1_btn.configure(state='normal')
        if self.fe2.dir_box.get("1.0", "end").strip() == const.strNoFileDisp:
            self.pick_asset.configure(values=('N/A'))
            self.asset_dict = []

    def update_list(self):
            if isinstance(self.asset_dict, dict) and self.fe1.dir_box.get("1.0", "end").strip() != const.strNoFileDisp:
                self.fe2.browse1_btn.configure(state='disabled')
                asset_list = tuple(self.asset_dict.keys())
            else:
                try:
                    file = self.fe2.dir_box.get('1.0','end').strip()
                    directory = self.dict_config[const.strWorkingFolder] + file
                    file_name = file.split('/')[-1]
                    try:
                        df = pd.read_excel(directory, usecols=[0,1,6])
                        df = df.loc[(df['EVENTS']==True) & (df['ASSET_TYPE']!=const.strAssetGen)]
                    except BaseException as err:
                        print(f'Error reading Asset file: [Errno] {err}\n Defaulting')
                        df = pd.read_excel(directory, usecols=[0,1])
                    df = df.sort_values(by='ASSET_ID')

                    try:
                        pa.detailed_validate_assets(df, file=file_name)
                    except FileNotFoundError as err:
                        messagebox.showwarning('WARNING', f'Validation was not possible:\n\n{err}')

                    self.asset_dict = dict(zip(df['ASSET_ID'], df['ASSET_TYPE']))
                    asset_list = tuple(self.asset_dict.keys())
                    if len(asset_list) == 0:
                        raise ValueError('No Relevant Assets in Asset File')
                except ValueError as err:
                    asset_list = ('N/A')
                    messagebox.showwarning('Warning',err)
                except BaseException as err:
                    asset_list = ('N/A')
                    print(err)

            return asset_list

    def add_zero(self, value):
        return value if int(value) > 9 else f'0{value}'

    def update_outage(self):
        file = self.dict_config[const.strWorkingFolder] + self.fe1.dir_box.get('1.0', 'end')[:-1]

        df = pd.read_excel(file, converters={'START_OUTAGE': pd.to_datetime,
                                             'END_OUTAGE': pd.to_datetime})

        self.df_backup = df.copy()

        if df.shape[1] != 7:
            raise TypeError('File Shape Error, aborting')

        delta_days = int(self.de3.get() )
        delta_hours = int(self.te3.hour.get())
        delta_minutes = int(self.te3.min.get())

        df['START_OUTAGE'] = df['START_OUTAGE'] + timedelta(days=delta_days, hours=delta_hours, minutes=delta_minutes)
        df['END_OUTAGE'] = df['END_OUTAGE'] + timedelta(days=delta_days, hours=delta_hours, minutes=delta_minutes)

        try:
            pa.auto_expand_and_save(df, file, index=False, sheet_name=const.strDataSheet)
            self.refresh_image()
            print(f'Saved to: {file}')
        except:
            raise TypeError('File Save Error')
        self.v = 1

        return 'update_outage'

    def restore_file(self):
        if isinstance(self.df_backup, pd.DataFrame) and self.v > 0:
            df = self.df_backup
            file = file = self.dict_config[const.strWorkingFolder] + self.fe1.dir_box.get('1.0', 'end')[:-1]
            try:
                pa.auto_expand_and_save(df, file, index=False, sheet_name=const.strDataSheet)
                self.refresh_image()
                print(f'Saved to: {file}')
            except PermissionError:
                messagebox.showinfo('Info', f'PermissionError: Close The Excel File and Try Again')
                print('File Save Error')

            self.v = 0
        else:
            raise TypeError('Nothing to Undo')

    def add_outage(self):
        start_date = datetime.strftime(self.de1.get_date(), '%Y-%m-%d') + f' {self.add_zero(self.te1.hour.get())}:{self.add_zero(self.te1.min.get())}:00'
        end_date = datetime.strftime(self.de2.get_date(), '%Y-%m-%d') + f' {self.add_zero(self.te2.hour.get())}:{self.add_zero(self.te2.min.get())}:00'

        days_difference = (datetime.strptime(end_date, '%Y-%m-%d %H:%M:%S') - datetime.strptime(start_date, '%Y-%m-%d %H:%M:%S')).days
        minutes_difference = (datetime.strptime(end_date[-8:], '%H:%M:%S') - datetime.strptime(start_date[-8:], '%H:%M:%S')).seconds / 86400
        difference = days_difference + minutes_difference
        if difference <= 0:
            messagebox.showinfo('Info', f'Enter a Future Date')
            raise TypeError('Enter a Future Date') 

        asset_id = self.pick_asset.get()
        if asset_id == '':
            raise TypeError('Select an Asset')

        if self.asset_dict == {}:
            raise TypeError('Select an Asset File')
        asset_type = self.asset_dict[asset_id]


        description = self.description_entry.get('1.0','end')[:-1]

        file = self.dict_config[const.strWorkingFolder] + self.fe1.dir_box.get('1.0', 'end')[:-1]

        if '.xlsx' not in file:
            folder = f'{self.dict_config[const.strWorkingFolder]}{const.strDataInput}'
            if self.checkbox_var1.get() == 'MAINT':
                file = folder + f'{datetime.strftime(self.ct, "%d-%m-%y")} {const.strMaintFile}_0.xlsx'
            else:
                file = folder + f'{datetime.strftime(self.ct, "%d-%m-%y")} {const.strContFile}_0.xlsx'
            i = 0
            while os.path.isfile(file):
                i += 1
                file = f'{file[:-7]}_{i}.xlsx'
            if not messagebox.askokcancel('Info', f'This will generate a new file in:\n{file}'):
                raise TypeError('Addition Cancelled')
                
            df = pd.DataFrame(columns=['ASSET_LONG_NAME','ASSET_ID','ASSET_TYPE','START_OUTAGE','END_OUTAGE','DURATION','DESCRIPTION'])

            self.fe1.dir_box['state'] = 'normal'
            self.fe1.dir_box.delete('1.0', 'end')
            self.fe1.dir_box.insert('1.0',file[len(self.dict_config[const.strWorkingFolder]):])
            self.fe1.dir_box['state'] = 'disabled'
            self.fe1.editdata1_btn['state'] = 'normal'

            self.maint_checkbox1['state'] = 'disabled'
            self.maint_checkbox2['state'] = 'disabled'
        
        else:
            df = pd.read_excel(file, converters={'START_OUTAGE': pd.to_datetime,
                                              'END_OUTAGE': pd.to_datetime})
            if df.shape[1] != 7:
                raise TypeError('File Shape Error, aborting')
                
        df.loc[len(df)] = [asset_id, asset_id, asset_type, start_date, end_date, difference, description]
        try:
            pa.auto_expand_and_save(df, file, index=False, sheet_name=const.strDataSheet)
            self.refresh_image()
            print(f'Saved to: {file}')
        except BaseException as err:
            print(err)
            self.clear_canvas()

        self.i += 1
        
        return 'add outage'

    def refresh_image(self, event=None):
        file = self.dict_config[const.strWorkingFolder] + self.fe1.dir_box.get('1.0', 'end')[:-1]
        out_type = 'MAINT' if self.checkbox_var1.get() else 'CONT'
        days = self.days_no.get()
        try:
            self.plot = Image.open(pa.plot_gantt_outages('', directory=file, no_of_days=days, save=True, out_type=out_type)).resize((self.cx,self.cy))
            self.pic_plot = ImageTk.PhotoImage(self.plot)
            self.plot_label.create_image(self.cx/2,self.cy/2,image=self.pic_plot)
        except FileNotFoundError as err:
            print(err)
        except BaseException as err:
            print(err)
            self.clear_canvas()
        return 'refresh_image'

    def undo(self):
        if self.i > 0:
            file = self.dict_config[const.strWorkingFolder] + self.fe1.dir_box.get('1.0', 'end')[:-1]

            df = pd.read_excel(file, converters={'START_OUTAGE': pd.to_datetime,
                                              'END_OUTAGE': pd.to_datetime})  
            if df.shape[1] != 7:
                raise TypeError('File Shape Error')
            
            df.drop(index=len(df)-1, inplace=True)
            try:
                pa.auto_expand_and_save(df, file, index=False, sheet_name=const.strDataSheet)
                self.refresh_image()
                print(f'Saved to: {file}')
            except BaseException as err:
                self.clear_canvas()
                print(err)
            self.i -= 1
            print('Addition Undone')
        else:
            raise TypeError('Nothing to Undo')
        return 'undo'
    
    def quit_me(self):
        print('Exiting')
        self.quit()
        self.destroy()
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------




# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
if __name__ == '__main__':
    App().mainloop()