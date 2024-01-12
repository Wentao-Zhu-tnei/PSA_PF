import tkinter as tk
import pandas as pd
import PSA_SND_Constants as const
import PSA_SND_Utilities as ut
import PSA_File_Validation as fv
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
    def __init__(self, parent, type='file'):
        super().__init__(parent)
        self.dict_config = ut.PSA_SND_read_config(const.strConfigFile)
        self.pic_excel_logo = tk.PhotoImage(file=self.dict_config[const.strWorkingFolder] + "/tkinterapp/assets/Excel_Logo.png").subsample(120)

        if type=='file':
            self.browse1_btn = tk.Button(self, text="Browse", command=lambda: self.fileBtn_clicked_TB(self.dir_box, '', self.editdata1_btn, True))
            self.editdata1_btn = ttk.Button(self, text="Edit Data", command=lambda: ut.open_xl(self.dir_box, self.dict_config[const.strWorkingFolder]), image=self.pic_excel_logo, compound=tk.LEFT, state=tk.DISABLED)
            self.editdata1_btn.grid(column=2, row=1, padx=10)
            width = 60
        elif type=='folder':
            self.browse1_btn = tk.Button(self, text="Browse", command=lambda: self.folderBtn_clicked(self.dir_box, '', False))
            width = 75
        else:
            raise ValueError('Invalid type')

        self.browse1_btn.grid(column=0, row=1, padx=10, ipadx=20)

        self.dir_box = tk.Text(self, width=width, height=1)
        self.dir_box.insert('1.0', const.strNoFileDisp)
        self.dir_box['state'] = 'disabled'
        self.dir_box.grid(column=1, row=1, padx=10)

    def fileBtn_clicked_TB(self, fLabel, fpat, fBtn, bInput):
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
        if bool(fBtn):
            if (fLabel.get("1.0", "end").strip() != const.strNoFileDisp):
                fBtn['state'] = 'normal'
            else:
                fBtn['state'] = 'disabled'

    def folderBtn_clicked(self, fLabel, fpat, bInput):
        fLabel['state'] = 'normal'
        fLabel.delete("1.0", "end")
        fLabel.insert("1.0", const.strNoFileDisp)
        #fLabel['state'] = 'disabled'
    
        if bInput:
            file_loc = self.dict_config[const.strWorkingFolder] + const.strDataInput
        else:
            file_loc = self.dict_config[const.strWorkingFolder] + const.strDataResults

        fldr = fd.askdirectory(title="Select Asset Data", initialdir=file_loc)
        fldr_extnd = fldr.replace(self.dict_config[const.strWorkingFolder], '')
        if fldr != '':
            #fLabel['state'] = 'normal'
            fLabel.delete("1.0", "end")
            fLabel.insert("1.0", fldr_extnd)
            fLabel['state'] = 'disabled'


# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Main App'''
class App(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        # -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '''Data Entry Parameters'''
        self.title('PSA Switch Generator')
        self.protocol("WM_DELETE_WINDOW", self.quit_me) 
        # Because of an issue with matplotlib, the program would actually keep running after the exit button is closed while a graph is displayed
        # To counter this, we set protocol to both quit and destroy when the exit button is pressed. 

        width = self.winfo_screenwidth() 
        height = self.winfo_screenheight()
        #self.state('normal') #normal, iconic, withdrawn, or zoomed
        self.geometry(f'1520x640+{(width-1500)//2}+{(height-620)//2}')
        
        self.dict_config = ut.PSA_SND_read_config(const.strConfigFile)
        self.df_nerda = []
        asset_value = tk.StringVar()
        self.status_value = tk.StringVar()

        self.ct = datetime.now()
        
        self.running = True
        self.df_backup = []
        self.i = 0 # Used for add service undo function
        self.v = 0 # Used for update services undo function

        self.bind('<Shift_L>', self.start_motor)
        self.bind('<KeyRelease-Shift_L>', self.stop_motor)
        #self.bind('<KeyPress-F5>', self.refresh_image) # Doesn't work because two arguments given when only one is taken (self, event) vs (self)

        self.LabelFrame00 = ttk.LabelFrame(self, text='Add New Event') # Data Entry
        frame6 = tk.Frame(self)
        self.LabelFrame10 = ttk.LabelFrame(frame6, text='Resulting Graph') # Plot Show
        self.LabelFrame20 = ttk.LabelFrame(frame6, text='Update Existing Switches')
        # -----------------------------------------------------------
        '''Frame 0: Select Maint File'''
        self.frame0 = ttk.LabelFrame(self, text='Select Switches File')
    
        self.fe1 = ExcelEntry(self.frame0)
        self.bind('<FocusOut>', self.clear_canvas)

        self.fe1.pack(padx=10, pady=10)
        # -----------------------------------------------------------
        self.frame001 = ttk.LabelFrame(self.LabelFrame00, text='Select Asset Data File')

        self.fe2 = ExcelEntry(self.frame001, type='file')
        self.fe2.pack(padx=10, pady=10)
        # -----------------------------------------------------------
        '''Frame 1: Select Asset, Enter Description'''
        frame1 = ttk.Frame(self.LabelFrame00)

        self.select_asset = tk.Label(frame1, text='Select Asset')
        self.select_asset.pack(padx=10, pady=5)

        self.asset_list = ('No_Asset_File')

        self.pick_asset = ttk.Combobox(frame1, values=self.asset_list, textvariable=asset_value, width = 35, postcommand=lambda: self.pick_asset.configure(values=self.update_list()))
        self.pick_asset.bind('<<ComboboxSelected>>', self.get_nerda)
        self.pick_asset.pack(padx=10, pady=0)

        self.status_label = tk.Label(frame1, text='')
        self.status_label.pack(padx=10, pady=8)

        self.description = tk.Label(frame1, text='Reason For Switch')
        self.description.pack(padx=10, pady=5)

        self.description_entry = tk.Text(frame1, height=3)
        self.description_entry.insert('1.0', 'Switching event')
        self.description_entry.pack(padx=10, pady=0)
        
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
        '''Input Change to status'''
        self.frame01 = tk.Frame(self.LabelFrame00)

        self.change_label = tk.Label(self.frame01, text='Change to:')
        self.change_label.grid(row=0, column=0, padx=5, pady=5)

        self.status_list = ('OPEN','CLOSED')
        self.pick_status = ttk.Combobox(self.frame01, values=self.status_list, textvariable=self.status_value, width = 15)
        self.pick_status.current(0)
        self.pick_status.grid(row=1, column=0, padx=5, pady=5)

        # -----------------------------------------------------------
        '''Input Scenario'''
        self.add_b = tk.StringVar(value='BASE')
        self.add_m = tk.StringVar(value='MAINT')
        self.add_c = tk.StringVar(value='CONT')
        self.add_mc = tk.StringVar(value='MAINT_CONT')

        parameters = (('Base', self.add_b)
                      ,('Maint', self.add_m)
                      ,('Cont', self.add_c)
                      ,('Maint_Cont', self.add_mc))
        
        self.scenario_label = tk.Label(self.frame01, text='In Scenario(s):')
        self.scenario_label.grid(row=0, column=1, columnspan=len(parameters), padx=5, pady=5)

        i = 1
        for params in parameters:
            text = params[0]
            variable = params[1]
            onvalue = text.upper()
            self.chk = tk.Checkbutton(self.frame01, text=text, variable=variable, onvalue=onvalue, command=lambda: self.scenario_check(checkbutton=self.chk))
            self.chk.grid(row=1, column=i, padx=5, pady=5, sticky='W')
            i += 1

        # -----------------------------------------------------------
        '''Frame 5: Update Existing Switches'''
        frame5 = tk.Frame(self.LabelFrame20)

        self.update_description = tk.Label(frame5, text='Add the following timedelta to all\nservices in the Switches File')

        self.delta_t = tk.Label(frame5, text='Days | Hours | Mins')
        self.te3 = HourEntry(frame5)
        self.te3.hourstr.set('0')

        self.daystr = tk.StringVar(self,'7')
        self.de3 = tk.Spinbox(frame5,from_=0,to=364,wrap=True,textvariable=self.daystr,width=2,font=('TkDefaultFont', 12),state="readonly")

        self.update_service_btn = tk.Button(frame5, text='Update Switches', command=self.update_service, bg='green', fg='white')
        self.update_undo_btn = tk.Button(frame5, text='Undo', command=self.restore_file, bg='black', fg='white')

        self.update_description.grid(padx=30, pady=10, column=0, row=0, rowspan=2)
        self.delta_t.grid(pady=10, column=1, row=0, columnspan=2)
        self.te3.grid(column=2, row=1)
        self.de3.grid(column=1, row=1)
        self.update_service_btn.grid(padx=30, pady=10, column=3, row=0, rowspan=2, sticky=tk.E)
        self.update_undo_btn.grid(pady=0, column=4, row=0, rowspan=2, sticky=tk.W)

        frame5.grid(column=0, row=0, padx=10, pady=10, sticky=tk.N)

        # -----------------------------------------------------------
        '''Grid Packing And Final Buttons for self.LabelFrame00'''
        self.frame001.grid(column=0, row=0, columnspan=2, padx=10, pady=0)
        self.frame0.grid(column=0, row=1, columnspan=2, padx=10, pady=0)
        self.frame01.grid(column=0, row=4, columnspan=2, padx=10, pady=2)
        frame1.grid(column=0, row=2, columnspan=2, padx=10, pady=0)
        frame4.grid(column=0, row=3, columnspan=2, padx=10, pady=10)

        self.add_service_btn = tk.Button(self.LabelFrame00, text='Add Event', command=self.add_service, bg='green', fg='white')
        self.add_service_btn.grid(pady=10, column=0, row=5, sticky=tk.E) 
        self.undo_btn = tk.Button(self.LabelFrame00, text='Undo', command=self.undo, bg='black', fg='white')
        self.undo_btn.grid(pady=10, column=1, row=5, sticky=tk.W)

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
        self.LabelFrame20.grid(padx=10, pady=0, column=0, row=1, sticky=tk.NE) # Update Switches
        self.LabelFrame10.grid(padx=10, pady=20, column=0, row=2, sticky=tk.NE) # Plot Show

        self.frame0.grid(padx=10, pady=10, column=0, row=0, columnspan=2, sticky=tk.S)
        frame6.grid(padx=10, pady=0, column=0, row=1, sticky=tk.NW)
        self.LabelFrame00.grid(padx=10, pady=0, column=1, row=1, sticky=tk.NW) # Add New Event
        
      
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
    def scenario_check(self, event=None, checkbutton=None):
        i = 0
        variables = (self.add_b, self.add_m, self.add_c, self.add_mc)
        for variable in variables:
            i += (variable.get() != '0')
        print(i)
        if i == 0:
            self.add_b.set('BASE')

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

        self.LabelFrame20.grid(padx=10, pady=0, column=0, row=1, sticky=tk.NE) # Update Outages
        self.LabelFrame10.grid(padx=10, pady=20, column=0, row=2, sticky=tk.NE) # Plot Show
        self.frame0.grid(padx=10, pady=10, column=0, row=0, columnspan=2, sticky=tk.S) # Select Maint file
        self.LabelFrame00.grid(padx=10, pady=10, column=1, row=1, sticky=tk.NW) # Add New Outage
        return

    def clear_canvas(self, event=None):
        if self.fe1.dir_box.get("1.0", "end").strip() == const.strNoFileDisp:
            self.plot_label.delete('all')
            self.fe2.browse1_btn.configure(state='normal')
        if self.fe2.dir_box.get("1.0", "end").strip() == const.strNoFileDisp:
            self.pick_asset.configure(values=('N/A'))

    def update_list(self):
        if isinstance(self.df_nerda, pd.DataFrame) and self.fe1.dir_box.get("1.0", "end").strip() != const.strNoFileDisp:
            self.fe2.browse1_btn.configure(state='disabled')
            asset_list = tuple(self.df_nerda.index)
        else:
            try:
                file = self.fe2.dir_box.get('1.0','end').strip()
                directory = self.dict_config[const.strWorkingFolder] + file
                file_name = file.split('/')[-1]
                df_asset = pd.read_excel(directory, sheet_name='NeRDA', usecols=[1,2,3,4])
                df_asset = df_asset[df_asset['SWITCHING']==True]
                df_asset = df_asset.set_index('NERDA_ID')

                try:
                    pa.detailed_validate_assets(df_asset, file=file_name)
                except FileNotFoundError as err:
                    messagebox.showwarning('WARNING', f'Validation was not possible:\n\n{err}')
            
                self.df_nerda = df_asset.copy()
                asset_list = tuple(df_asset.index)
                if len(asset_list) == 0:
                    raise ValueError('No Relevant Switches in Switching File')
            except ValueError as err:
                asset_list = ('N/A')
                messagebox.showwarning('Warning',err)
            except BaseException as err:
                asset_list = ('N/A')
                print(err)
        return asset_list

    def get_nerda(self, event=None):
        asset = self.pick_asset.get()
        status = self.df_nerda.loc[asset,'DEFAULT_PF_POSITION']
        self.status_label.configure(text=f'The default PF position of switch: {self.pick_asset.get()} is {status}')
        if status == 'OPEN':
            self.status_value.set('CLOSED')
        elif status == 'CLOSED':
            self.status_value.set('OPEN')
        else:
            self.status_value.set('OPEN') # If Status is UNKNOWN
        return status

    def add_zero(self, value):
        return value if int(value) > 9 else f'0{value}'

    def update_service(self):
        file = self.dict_config[const.strWorkingFolder] + self.fe1.dir_box.get('1.0', 'end')[:-1]

        df = pd.read_excel(file, converters={'START_SERVICE': pd.to_datetime,
                                             'END_SERVICE': pd.to_datetime})

        self.df_backup = df.copy()

        if df.shape[1] != 9:
            raise TypeError('File Shape Error, aborting')

        delta_days = int(self.de3.get() )
        delta_hours = int(self.te3.hour.get())
        delta_minutes = int(self.te3.min.get())

        df['START_SERVICE'] = df['START_SERVICE'] + timedelta(days=delta_days, hours=delta_hours, minutes=delta_minutes)
        df['END_SERVICE'] = df['END_SERVICE'] + timedelta(days=delta_days, hours=delta_hours, minutes=delta_minutes)

        try:
            pa.auto_expand_and_save(df, file, index=False, sheet_name=const.strDataSheet)
            self.refresh_image()
            print(f'Saved to: {file}')
        except BaseException as err:
            print(err)
            self.clear_canvas()
        self.v = 1

        return 'update_service'

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

    def add_service(self):
        start_date = datetime.strftime(self.de1.get_date(), '%Y-%m-%d') + f' {self.add_zero(self.te1.hour.get())}:{self.add_zero(self.te1.min.get())}:00'
        end_date = datetime.strftime(self.de2.get_date(), '%Y-%m-%d') + f' {self.add_zero(self.te2.hour.get())}:{self.add_zero(self.te2.min.get())}:00'

        days_difference = (datetime.strptime(end_date, '%Y-%m-%d %H:%M:%S') - datetime.strptime(start_date, '%Y-%m-%d %H:%M:%S')).days
        minutes_difference = (datetime.strptime(end_date[-8:], '%H:%M:%S') - datetime.strptime(start_date[-8:], '%H:%M:%S')).seconds / 86400
        difference = days_difference + minutes_difference
        if difference <= 0:
            messagebox.showinfo('Info', f'Enter a Future Date')
            raise TypeError('Enter a Future Date') 

        asset_id = self.pick_asset.get()
        if asset_id == '' or asset_id == 'N/A':
            raise TypeError('Select an Asset')

        status = self.status_value.get()

        asset_type = self.df_nerda.loc[asset_id,'ASSET_TYPE']

        description = self.description_entry.get('1.0','end')[:-1]

        file = self.dict_config[const.strWorkingFolder] + self.fe1.dir_box.get('1.0', 'end')[:-1]

        if '.xlsx' not in file:
            folder = f'{self.dict_config[const.strWorkingFolder]}{const.strDataInput}'
            file = folder + f'{datetime.strftime(self.ct, "%d-%m-%y")} {const.strSwitchFile}_0.xlsx'
            i = 0
            while os.path.isfile(file):
                i += 1
                file = f'{file[:-7]}_{i}.xlsx'
            if not messagebox.askokcancel('Info', f'This will generate a new file in:\n{file}'):
                raise TypeError('Addition Cancelled')
                
            df = pd.DataFrame(columns=['ASSET_LONG_NAME','ASSET_ID','ASSET_TYPE','START_SERVICE','END_SERVICE','DURATION','ACTION','DESCRIPTION','SCENARIO'])

            self.fe1.dir_box['state'] = 'normal'
            self.fe1.dir_box.delete('1.0', 'end')
            self.fe1.dir_box.insert('1.0',file[len(self.dict_config[const.strWorkingFolder]):])
            self.fe1.dir_box['state'] = 'disabled'
            self.fe1.editdata1_btn['state'] = 'normal'
        
        else:
            df = pd.read_excel(file, converters={'START_SERVICE': pd.to_datetime,
                                              'END_SERVICE': pd.to_datetime})
            if df.shape[1] != 9:
                raise TypeError('File Shape Error, aborting')

        for scenario in [self.add_b.get(), self.add_m.get(), self.add_c.get(), self.add_mc.get()]:
            if scenario != '0':
                df.loc[len(df)] = [asset_id, asset_id, asset_type, start_date, end_date, difference, status, description, scenario]
        try:
            pa.auto_expand_and_save(df, file, index=False, sheet_name=const.strDataSheet)
            self.refresh_image()
            print(f'Saved to: {file}')
        except BaseException as err:
            print(err)
            self.clear_canvas()

        self.i += 1
        return 'add service'

    def refresh_image(self, event=None):
        file = self.dict_config[const.strWorkingFolder] + self.fe1.dir_box.get('1.0', 'end')[:-1]
        days = self.days_no.get()
        try:
            self.plot = Image.open(pa.plot_gantt_outages('', directory=file, no_of_days=days, save=True, out_type='SWITCH')).resize((self.cx,self.cy))
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

            df = pd.read_excel(file, converters={'START_SERVICE': pd.to_datetime,
                                              'END_SERVICE': pd.to_datetime})  
            if df.shape[1] != 9:
                raise TypeError('File Shape Error')
            
            df.drop(index=len(df)-1, inplace=True)
            try:
                pa.auto_expand_and_save(df, file, index=False, sheet_name=const.strDataSheet)
                self.refresh_image()
                print(f'Saved to: {file}')
            except BaseException as err:
                self.plot_label.delete('all')
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