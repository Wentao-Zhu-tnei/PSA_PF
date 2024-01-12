import tkinter as tk
import PSA_SND_Constants as const
import PSA_SND_Utilities as ut
import PSA_analysis as pa
import os

from tkinter import ttk, messagebox
from tkinter import filedialog as fd

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Main App'''
class App(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        # -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '''Config'''
        self.title('Inter Run Analysis')
        self.protocol("WM_DELETE_WINDOW", self.quit_me)         
        self.dict_config = ut.PSA_SND_read_config(const.strConfigFile)

        # -----------------------------------------------------------
        '''Input Parameters''' 
        LabelFrame00 = ttk.LabelFrame(self, text='Input Parameters')
        LabelFrame00.columnconfigure(0, weight=1)
        LabelFrame00.columnconfigure(1, weight=1)
        LabelFrame00.columnconfigure(2, weight=1)

        LabelFrame01 = ttk.LabelFrame(self, text='Data/Results Reference')

        # Ref_dir
        self.directory = f'{self.dict_config[const.strWorkingFolder]}{const.strDataResults}'
        self.fname1 = tk.Text(LabelFrame01, width=80, height=1)
        self.fname1.insert("1.0", self.directory)
        self.fname1['state'] = 'disabled'
        self.fname1.pack(side=tk.RIGHT, padx=10, pady=10)
        self.fldrBtn = ttk.Button(LabelFrame01, text="Select Folder for Analysis", command=lambda: self.folderBtn_clicked(self.fname1, False))
        self.fldrBtn.pack(side=tk.LEFT, padx=10, pady=10)

        temp_directory = os.listdir(self.directory)
        self.runs_list = tuple(pa.get_applicable_runs(directory=self.directory, target='AUT')[:200])
        
        # Start
        self.start_label = tk.Label(LabelFrame00, text='Start')
        self.start_label.grid(row=0, column=0, sticky='N', padx=5, pady=5)

        self.PSA_start = tk.StringVar(value=self.runs_list[0])
        self.cboxstart = ttk.Combobox(LabelFrame00, values=self.runs_list, textvariable=self.PSA_start, width = 35)
        self.cboxstart.bind('<<ComboboxSelected>>', self.update_stop)
        self.cboxstart.bind('<<ComboboxSelected>>', self.update_assets, add='+')
        self.cboxstart.bind('<<ComboboxSelected>>', self.update_scenarios, add='+')
        self.cboxstart.grid(row=1, column=0, sticky='N', padx=5, pady=5)

        # Stop
        self.stop_runs = pa.limit_runs(self.runs_list, start=self.PSA_start.get())
        self.stop_label = tk.Label(LabelFrame00, text='Stop (Included)')
        self.stop_label.grid(row=0, column=1, sticky='N', padx=5, pady=5)

        self.PSA_end = tk.StringVar(value=self.stop_runs[-1])
        self.cboxend = ttk.Combobox(LabelFrame00, values=self.stop_runs, textvariable=self.PSA_end, width = 35)
        self.cboxend.grid(row=1, column=1, sticky='N', padx=5, pady=5)

        # Step
        self.step_label = tk.Label(LabelFrame00, text='Step (Every nth Run)')
        self.step_label.grid(row=0, column=2, sticky='N', padx=5, pady=5)

        self.PSA_step = tk.IntVar(value=16)
        self.textbox = ttk.Entry(LabelFrame00, textvariable=self.PSA_step, width=35)
        self.textbox.grid(row=1, column=2, sticky='N', padx=5, pady=5)

        # Daily Flag
        self.daily = tk.BooleanVar(value=False)
        self.daily_chk = ttk.Checkbutton(LabelFrame00, variable=self.daily, text='Take One Run From Each Day?', command=self.toggle_step)
        self.daily_chk.grid(row=2, column=2)

        # -----------------------------------------------------------
        '''Analysis Functions''' 
        LabelFrame10 = ttk.LabelFrame(self, text='Analysis Functions')
        LabelFrame10.columnconfigure(0, weight=1)
        LabelFrame10.columnconfigure(1, weight=4)
        LabelFrame10.columnconfigure(2, weight=1)

        LabelFrame11 = ttk.LabelFrame(LabelFrame10, text='')
        LabelFrame11.columnconfigure(0, weight=1)
        LabelFrame11.columnconfigure(1, weight=1)

        # Dropdown for Scenario
        self.scenario_value = tk.StringVar(value='Undefined')
        self.pick_scenario = ttk.Combobox(LabelFrame10, values=('Undefined'), textvariable=self.scenario_value, width = 15)
        self.pick_scenario.bind('<<ComboboxSelected>>', self.update_assets)
        self.pick_scenario.grid(row=3, column=0, padx=10, pady=5, sticky=tk.NW)
        self.update_scenarios()

        # Dropdown for Asset
        self.asset_value = tk.StringVar(value='Undefined')
        self.pick_asset = ttk.Combobox(LabelFrame10, values=('Undefined'), textvariable=self.asset_value, width = 25)
        self.pick_asset.grid(row=1, column=0, padx=10, pady=5, sticky=tk.NW)    
        self.update_assets()

        # Labels
        #self.analysis_label = tk.Label(LabelFrame11, text='Analysis Functions')
        #self.analysis_label.grid(row=0, column=0, columnspan=2, padx=5, pady=5)
        self.asset_label = tk.Label(LabelFrame10, text='Select AssetID to Analyse')
        self.asset_label.grid(row=0, column=0, padx=10, pady=0, sticky=tk.SW)
        self.scenario_label = tk.Label(LabelFrame10, text='Select Scenario from Dropdown')
        self.scenario_label.grid(row=2, column=0, padx=10, pady=0, sticky=tk.SW)
        
        # Variables
        self.selected_function = tk.StringVar(value='Read HL Overview')
  
        args = (('Plot SIA',1,0),
                ('Plot Asset',2,0),
                ('Read NeRDA',3,0),
                ('Read HL Overview',4,0))

        for arg in args:
            self.chk = ttk.Radiobutton(LabelFrame11, text=arg[0], value=arg[0], variable=self.selected_function)
            self.chk.grid(row=arg[1], column=arg[2], padx=10, pady=5, sticky=tk.W)

        # Display Results Button
        self.display_btn = ttk.Button(LabelFrame10, text='Display Results', command=self.display_results)
        self.display_btn.grid(row=0, column=2, padx=10, pady=5, sticky=tk.SE)

        # -----------------------------------------------------------
        '''Final Packing''' 
        LabelFrame01.grid(padx=15, pady=5, column=0, row=0, sticky=tk.EW)
        LabelFrame00.grid(padx=15, pady=5, column=0, row=1, sticky=tk.EW)
        LabelFrame11.grid(row=0, rowspan=4, column=1, sticky=tk.EW, pady=5)
        LabelFrame10.grid(padx=15, pady=5, column=0, row=2, sticky=tk.EW)

    # -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    '''Functions for input processing'''
    def display_results(self):
        directory = self.directory
        func = self.selected_function.get()
        runs_list = self.stop_runs
        stop = self.PSA_end.get()
        step = int(self.PSA_step.get())
        runs_dict = dict(zip(runs_list,list(range(0,len(runs_list)))))
        flag = bool(self.daily.get())
        if flag:
            step = 1
        try:
            runs = runs_list[0:runs_dict[stop]+1:step]
            first = runs[0]
            runs = pa.limit_runs(runs, first, daily_flag=flag)
            #print(runs)
        except IndexError as err:
            print(err)
            runs = []

        if len(runs) == 1 and len(runs_list[0:runs_dict[stop]+1]) != 1:
            if not messagebox.askyesno('One Run', f'With the current start/stop/step setup, you are only analysing one run: {runs[0]}.\
                                \n\nAre you sure you want to proceed with the current settings?'):
                print('Process Cancelled by user')
                return
        elif len(runs) > 20:
            if not messagebox.askyesno('Lots of Runs', f'With the current start/stop/step setup, you would analyse {len(runs)} runs. This may take some time and the legend may be hidden.\
                                \n\nAre you sure you want to proceed with the current settings'):
                print('Process Cancelled by user')
                return
        elif len(runs) == 0:
            messagebox.showerror('No Runs', 'With the current start/stop/step setup, no runs would be valid for analysis.\
                                 \n\nPlease try entering different parameters')
            return

        status = const.PSAok
        if func == 'Read HL Overview':
            try:
                filepath, vhl_dict = pa.VHL_overview(sorted_dir=runs, stopping_index=len(runs), directory=directory, target='AUT', save=True)
            except BaseException as err:
                status = err
        elif func == 'Plot SIA':
            primary = self.asset_value.get()
            try:
                filepath = pa.plot_inter_run_sia(directory, primary, start=runs, save=True)
            except BaseException as err:
                status = err
        elif func == 'Plot Asset':
            asset = self.asset_value.get()
            try:
                filepath = pa.plot_inter_run_flex_rq(directory, asset, start=runs, scenario='BASE', save=True)
            except BaseException as err:
                status = err
        elif func == 'Read NeRDA':
            try:
                filepath = pa.plot_inter_run_nerda(directory, start=runs, save=True)
            except BaseException as err:
                status = err
        else:
            raise KeyError('None of the checkboxes were checked!')
        if status == const.PSAok:
            os.system('"' + filepath + '"')
        else:
            print(status)

    def folderBtn_clicked(self, fLabel, bInput):
        fLabel['state'] = 'normal'
        fLabel.delete("1.0", "end")
        fLabel.insert("1.0", const.strNoFileDisp)
        fLabel['state'] = 'disabled'

        if bInput:
            file_loc = self.dict_config[const.strWorkingFolder] + const.strDataInput
        else:
            file_loc = self.dict_config[const.strWorkingFolder] + const.strDataResults

        fldr = fd.askdirectory(title="Select Data/Results Folder", initialdir=file_loc)
        if fldr != '':
            fLabel['state'] = 'normal'
            fLabel.delete("1.0", "end")
            fLabel.insert("1.0", fldr)
            fLabel['state'] = 'disabled'

            self.directory = fldr
            self.update_start()
            self.update_stop()
            self.update_scenarios()
            self.update_assets()

    def update_scenarios(self, event=None):
        directory = self.directory
        first = self.stop_runs[0]
        files = os.listdir(f'{directory}/{first}')
        scenarios = []
        for scenario in [const.strBfolder.strip('\\'), const.strMfolder.strip('\\'), const.strCfolder.strip('\\'), const.strMCfolder.strip('\\')]:
            if scenario in files:
                scenarios.append(scenario.split(' - ')[1])

        if scenarios == []:
            print('Error resolving scenarios')
            scenarios = ['BASE','MAINT','CONT','MAINT_CONT']
        
        self.pick_scenario.configure(values=scenarios)
        self.scenario_value.set(scenarios[0])

        return scenarios

    def update_assets(self, event=None):
        directory = self.directory
        scenario = self.scenario_value.get()
        scenario_dict = {'BASE': const.strBfolder,
                         'MAINT': const.strMfolder,
                         'CONT': const.strCfolder,
                         'MAINT_CONT': const.strMCfolder}
           
        assets, i = [], 0
        while assets == []:
            try:
                first = self.stop_runs[i]
            except IndexError:
                assets = ['No_Assets_Found']
            for asset_type in ['LN_LDG_DATA','TX_LDG_DATA']:
                try:
                    files = os.listdir(f'{directory}/{first}{scenario_dict[scenario]}{asset_type}')
                    assets += [asset.split('-')[1].removesuffix('.xlsx') for asset in files]
                except BaseException:
                    i += 1
                
        
        assets = tuple(assets)

        self.pick_asset.configure(values=assets)
        self.asset_value.set(assets[0])
    
        return assets    

    def update_stop(self, event=None):
        self.stop_runs = pa.limit_runs(self.runs_list,self.PSA_start.get())
        self.cboxend.configure(values=self.stop_runs)
        self.cboxend.current(len(self.stop_runs)-1)

    def update_start(self):
        self.runs_list = tuple(pa.get_applicable_runs(directory=self.directory, target='AUT')[:200])
        self.cboxstart.configure(values=self.runs_list)
        self.cboxstart.current(0)

    def toggle_step(self, event=None):
        if self.daily.get():
            self.textbox.configure(state='disabled')
        else:
            self.textbox.configure(state='normal')

    def quit_me(self):
        print('Exiting')
        self.quit()
        self.destroy()

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
if __name__ == '__main__':
    App().mainloop()