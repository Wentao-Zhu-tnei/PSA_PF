import os
import tkinter as tk
from tkinter import *
from tkinter import ttk, Text
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo, askyesno, showwarning
import pandas as pd
import PSA_GUI_Outputs as op
import PSA_SND_Utilities as ut
import PSA_File_Validation as fv
import PSA_SND_Constants as const
import PSA_Calc_Flex_Reqts as fq
import PF_Config as pfconf
from pathlib import Path
from datetime import datetime, timedelta
import time
import PSA_PF_SF as pfsf
import PSA_Constraint_Resolution as cr
import PSA_analysis as pa
import PSA_SND_File_creation as SNDfile
import PSA_ProgressBar as pbar
import subprocess as subp

start_time = datetime.now()
class App(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.title("PSA - Event Manager")
        ws = 800
        hs = 350
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw/2) - (ws/2)
        y = (sh/2) - (hs/2)
        wstr = str(ws) + "x" + str(hs) + "+" + str(int(x)) + "+" + str(int(y))
        self.geometry(wstr)
        self.resizable(0,0)

        Frame01 = tk.Frame(self)
        LabelFrame01 = ttk.LabelFrame(self, text='CORE FUNCTIONS')
        LabelFrame02 = ttk.LabelFrame(Frame01, text='OUTPUT FILE ANALYSIS')
        LabelFrame03 = ttk.LabelFrame(Frame01, text='INPUT FILE GENERATION')

        # Buttons for each event
        self.Constraints = tk.Button(LabelFrame01, text="Calculate Flexibility Requirements", command=self.__create_Constraints)
        self.Constraints.pack(padx=10, pady=10, ipadx=10, expand=True, fill=tk.X, side=tk.LEFT)
        
        self.SensitivityFactors = tk.Button(LabelFrame01, text="Sensitivity Factors",command=self.__calc_SF)
        self.SensitivityFactors.pack(padx=10, pady=10, ipadx=10, expand=True, fill=tk.X, side=tk.LEFT)
        
        self.AcceptableSolution = tk.Button(LabelFrame01, text="Constraint Resolution", command=self.__constraint_resolution)
        self.AcceptableSolution.pack(padx=10, pady=10, ipadx=10, expand=True, fill=tk.X, side=tk.LEFT)

        self.AnalysisnReports = tk.Button(LabelFrame02, text="Analysis & Reports", command=self.__analysis)
        self.AnalysisnReports.grid(row=0, column=0, padx=10, pady=10, ipadx=10, sticky=tk.EW)

        self.autorun_var = BooleanVar(value=True)
        self.AnalysisnReports_check = ttk.Checkbutton(LabelFrame02, text = "Autorun A&R?", variable=self.autorun_var)
        self.AnalysisnReports_check.grid(row=0, column=1, padx=10, pady=10, ipadx=10, sticky=tk.EW)

        self.InterAnalysis = tk.Button(LabelFrame02, text="Inter-Run Analysis", command=lambda: subp.Popen(['python','PSA_Inter_Run_Analysis_App.py']))
        self.InterAnalysis.grid(row=1, column=0, padx=10, pady=10, ipadx=10, sticky=tk.EW)
        
        self.SNDFiles = tk.Button(LabelFrame03, text="Create S&D Test Files", command=self.__createSNDfiles)
        self.SNDFiles.grid(padx=10, pady=10, ipadx=10, column=0, row=0, sticky=tk.NSEW)

        self.MaintFiles = tk.Button(LabelFrame03, text="Create Outage Files", command=lambda: subp.Popen(['python','PSA_Maint_Generator_App.py']))
        self.MaintFiles.grid(padx=10, pady=10, ipadx=10, column=1, row=0, sticky=tk.NSEW)

        self.AssetFiles = tk.Button(LabelFrame03, text="Choose Asset Thresholds", command=lambda: subp.Popen(['python','PSA_Asset_Threshold_App.py']))
        self.AssetFiles.grid(padx=10, pady=10, ipadx=10, column=0, row=1, sticky=tk.NSEW)
        
        self.EventFiles = tk.Button(LabelFrame03, text="Create Event Files", command=lambda: subp.Popen(['python','PSA_Event_Generator_App.py']))
        self.EventFiles.grid(padx=10, pady=10, ipadx=10, column=1, row=1, sticky=tk.NSEW)

        self.SwitchFiles = tk.Button(LabelFrame03, text="Create Switch Files", command=lambda: subp.Popen(['python','PSA_Switching_Event_App.py']))
        self.SwitchFiles.grid(padx=10, pady=10, ipadx=10, column=0, row=2, sticky=tk.NSEW)

        LabelFrame01.pack(padx=10, pady=10, ipadx=10, expand=True, fill=tk.X)
        LabelFrame03.pack(padx=10, pady=10, ipadx=10, expand=True, fill=tk.X, side=tk.LEFT)
        LabelFrame02.pack(padx=10, pady=10, ipadx=10, expand=True, fill=tk.X, side=tk.LEFT)
        Frame01.pack(padx=10, pady=10, ipadx=10, expand=True, fill=tk.X)

        self.Help = tk.Button(self, text="User Guide", bg="green", fg="white", command=self.helpBtn_clicked)
        self.Help.pack(padx=10, pady=10, ipadx=10, side=tk.LEFT)
        
        self.Exit = tk.Button(self, text="Exit", bg="black", fg="white", command=self.exitMainBtn_clicked)
        self.Exit.pack(padx=10, pady=10, ipadx=10, side=tk.RIGHT)

        self.dict_config = ut.PSA_SND_read_config(const.strConfigFile)
        strVersion = "1.1"
        msg = "********************************************************\n\n"
        #msg = msg + "PSA Version: \t" + strVersion + "\n\n"
        msg = msg + "Working Folder : \t" + self.dict_config[const.strWorkingFolder] + "\n\n"
        msg = msg + "PSA2SND Folder : \t" + self.dict_config[const.strFileDetectorFolder] + "\n\n"
        if pfconf.bDebug:
            msg = msg + "Debug Mode: \t" + "ON\n\n"
        else:
            msg = msg + "Debug Mode: \t" + "OFF\n\n"
        msg = msg + "********************************************************\n"
        title = "PSA Configuration"
        showinfo(title, msg, parent=self)

    def helpBtn_clicked(self):
        answer = askyesno("User Guide", message="Are you sure?", parent=self)
        if answer:
            if os.path.isfile(const.strUserGuide):
                os.system(const.strUserGuide)
            else:
                status = const.PSAfileExistError
                msg = "User Guide doesn't exist: " + const.strUserGuide
                showwarning(title="Error", message=msg, parent=self.root)

    # Button Def #
    def stopBtn_clicked(self):
        answer = askyesno("INFO", message="Are you sure you want to stop the process?", parent=self.root)
        if answer:
            if not self.bStopProcess:
                self.bStopProcess = True
                showinfo(title="INFO", message="Process has been terminated by user", parent=self.root)
                self.enableButtons(self.fileBtn1, self.fileBtn3, self.events, self.plan, self.unpl, self.switch, self.xlBtn3,
                                    self.view_outage_data_btn, self.view_threshold_data_btn,
                                    self.switchPFLag, self.switchPFLead, 
                                    self.entryPowerFactor, self.entryLneThresh, self.entryTrafoThresh,
                                    self.switchAutomatic, self.switchAutomaticSimplified, self.switchManual, 
                                    self.calBtn, self.entryAutoRuns, self.entryNumDays)
                self.varDisplayPSArunID.set('Process terminated by user')
                self.root.update_idletasks()

    def exitMainBtn_clicked(self):
        answer = askyesno("Exit Application", message="Are you sure?", parent=self)
        if answer:
            self.quit()
            self.destroy()
            exit()

    def exitBtn_clicked(self):
        answer = askyesno("Exit Window", message="Are you sure?", parent=self.root)
        if answer:
            #self.root.quit() # This causes the entire mainloop to stop, not just the toplevel window. Not sure why.
            self.root.destroy()
            return

    def disableButtons(self, *args):
        for button in args:
            button['state'] = 'disabled'
        for radiobutton in args:
            radiobutton['state'] = 'disabled'
        for checkbutton in args:
            checkbutton['state'] = 'disabled'

    def enableButtons(self, *args):
        for button in args:
            button['state'] = 'normal'
        for radiobutton in args:
            radiobutton['state'] = 'normal'
        for checkbutton in args:
            checkbutton['state'] = 'normal'

    def repeatAuto(self):
        # new calculation for repeat time to take into account of processing time            
        if not self.bStopProcess:
            self.count += 1
            if (self.NumAutoRuns == -1) or (self.count <= self.NumAutoRuns):
                if (self.NumAutoRuns == -1):
                    print("Automated mode until user terminates the process")
                else:
                    if self.count > 1:
                        pbar.delProgressBar()
                    print("Automated run cycle: " + str(self.count) + " of " + str(self.NumAutoRuns))
                print("Start time of: " + str(datetime.now()))
                t0 = time.time()
                PSAstatus, PSArunID = self.calcBtn_clicked_core()
                if PSAstatus == const.PSAok:
                    t1 = time.time()
                    t2 = int(t1 -t0)
                    self.sleep_time = int(self.dict_config[const.strCalcFlexReqtsInterval])
                    if t2 <= self.sleep_time:
                        self.sleep_time = self.sleep_time - t2
                    else:
                        self.sleep_time = 1
                    if (self.NumAutoRuns == -1) or (self.count < self.NumAutoRuns):
                        if pfconf.bDebug:
                            print("Calc time secs = " + str(t2))
                            print("Repeat time secs = " + str(self.dict_config[const.strCalcFlexReqtsInterval]))
                            print("Sleep time secs = " + str(self.sleep_time))
                        t2 = datetime.now() + timedelta(seconds=self.sleep_time)
                        print("Next start time = " + str(t2))
                        pbar.delProgressBar()
                        pbar.createProgressBar("Waiting to start next run", False)
                        pbar.updateProgressBar(0, "Next run starts at " + str(t2)[:19])
                        # convert secs to millisecs
                        self.sleep_time = self.sleep_time * 1000
                        self.root.after(self.sleep_time, self.repeatAuto)
                    else:
                        print("*** Automated runs completed after " + str(self.NumAutoRuns) + " cycles ***")
            else:
                print("*** Automated runs completed after " + str(self.NumAutoRuns) + " cycles ***")
        return PSAstatus, PSArunID

    def get_running_mode(self):
        self.mode = self.v.get()  # 1 - auto; 2 - simplified auto; 3 - manual
        if self.mode == 1:
            self.strMode = const.strRunModeAUT
            self.auto = True
        elif self.mode == 2:
            self.strMode = const.strRunModeSEMIAUT
            self.auto = True
        elif self.mode == 3:
            self.strMode = const.strRunModeMAN
            self.auto = False

    def calcBtn_clicked_core(self):
        start_time_flex_reqts = datetime.now()
        PSAstatus = const.PSAok
        lCombinedFlexReqts, lCombinedConstraints = ([] for i in range(2))
        dfCombinedEvents, dfCombinedConstraints = (pd.DataFrame() for i in range(2))
        self.BSPmodel = self.dict_config[const.strWorkingFolder] + self.fname1.get("1.0", END).strip()
        # self.primary = self.dict_config[const.strWorkingFolder] + self.fname2.get("1.0", END).strip()
        self.primary = self.dict_config[const.strWorkingFolder] + self.cboxBSPs.get()
        self.assetFile = self.dict_config[const.strWorkingFolder] + self.fname3.get("1.0", END).strip()
        self.bEvents = self.eventsvar.get()
        self.eventFile = self.dict_config[const.strWorkingFolder] + self.fname9.get("1.0", END).strip()
        self.plout = self.planvar.get()
        self.ploutFile = self.dict_config[const.strWorkingFolder] + self.fname4.get("1.0", END).strip()
        self.unplout = self.unplvar.get()
        self.unploutFile = self.dict_config[const.strWorkingFolder] + self.fname5.get("1.0", END).strip()
        self.bSwitch = self.switchvar.get()
        self.switchFile = self.dict_config[const.strWorkingFolder] + self.fname6.get("1.0", END).strip()

        #self.timeStep = self.cboxTimeSteps.get()
        self.timeStep = "NULL"
        #self.kWstep = self.cboxConstraintkWStep.get()
        self.kWstep = "NULL"
        self.powFactor = self.entryPowerFactor.get()
        self.lagging = self.varPF.get()  # Leading - False, Lagging -True
        self.output = self.outputvar.get()
        # self.SPM = self.SPMvar.get()
        # self.SEPM = self.SPEvar.get()
        # self.SEC = self.SECvar.get()
        # self.DYN = self.DYNvar.get()
        # self.auto = self.varAutoMan.get()

        self.LneThresh = self.varLneThresh.get()
        self.TrafoThresh = self.varTrafoThresh.get()
        self.NumDays = self.varNumDays.get()
        self.NumAutoRuns = self.varAutoRuns.get()
        PSA_running_mode = op.getPSARunningMode(self.plout, self.unplout)
        print("################################################################\n")
        # print(f'Running in AUTO mode?: [{self.auto}]')
        print(f'Running in [{self.strMode}] mode.')
        print(f'Running PSA scenario: [{PSA_running_mode}]')
        print("################################################################\n")

        [msg, bMsg, PSArunID, lModels, fnme_BSPmodel, fnme_primary, fnme_assetFile, fnme_eventFile, fnme_ploutFile, fnme_unploutFile, fnme_switchFile] \
            = op.PSA_create_output(const.strConfigFile, const.strWorkingFolder, self.BSPmodel, self.primary,
                                   self.assetFile, self.bEvents, self.eventFile, self.plout, self.ploutFile, self.unplout, self.unploutFile,
                                   self.bSwitch, self.switchFile, self.auto)

        if bMsg:
            PSAstatus = const.PSAfileReadError
            showwarning(title="WARNING", message=msg, parent=self.root)
        else:
            msg, bWarning = fv.validateFlexReqtsData(fnme_BSPmodel=fnme_BSPmodel,proj_dir=Path(self.BSPmodel).parent.absolute(),
                                                     refAssetFile=os.path.join(self.dict_config[const.strWorkingFolder],
                                                     const.strRefAssetFile),assetFile=self.assetFile, bEvent=self.bEvents,
                                                     eventFile=self.eventFile,plout=self.plout, ploutFile=self.ploutFile,
                                                     unplout=self.unplout,unploutFile=self.unploutFile,
                                                     bSwitch=self.bSwitch, switchFile=self.switchFile)
            if bWarning:
                PSAstatus = const.PSAfileReadError
                showwarning(title='WARNING', message=msg, parent=self.root)
            else:
                PSAfolder = op.makeResultsFolder(self.dict_config[const.strWorkingFolder] + const.strDataResults, PSArunID)
                self.varDisplayPSArunID.set("Processing: " + PSArunID)
                self.root.update_idletasks()

                dict_PSA_params = ut.create_dictPSAparams(fnme_BSPmodel, fnme_primary, fnme_assetFile, self.bEvents, fnme_eventFile, self.plout,
                                                          fnme_ploutFile, self.unplout, fnme_unploutFile,
                                                          self.bSwitch, fnme_switchFile,
                                                          self.timeStep, self.kWstep, self.powFactor, self.lagging,
                                                          self.LneThresh, self.TrafoThresh, self.auto, self.strMode,
                                                          self.NumDays)
                ut.writePSAparams(PSAfolder, PSArunID, dict_PSA_params)
 
                #calculate HH index from PSArunID
                hr = int(PSArunID.split("-")[3])
                min = int(PSArunID.split("-")[4])
                if min >= 30:
                    start_time_row = hr * 2 + 1
                else:
                    start_time_row = hr * 2

                ### MANUAL MODE ###
                lNamesGrid=['BERI','KENN','ROSH','WALL'] ### these are the grids which are used for Downscaling the Sia forecast
                print("################################################################")
                print(f'These are the grids that will be used for Downscaling the Sia forecast\n{lNamesGrid}')
                print("################################################################")
                if self.strMode == const.strRunModeMAN:
                    self.count = 1
                    self.NumAutoRuns = 1
                    strInputFolder, strResultsFolder, strNeRDAFolder, strSIAFolder = \
                        op.PSA_create_man_results_folders(PSAfolder,self.plout,self.unplout)
                    op.copy2ResultsFolder(strInputFolder, lModels)
                    status, msg, end_time_flex_reqts_manual, dictSIAFeeder, dictSIAGen, dictSIAGroup, start_time_base, \
                    NeRDA_output_path, dfDownScalingFactor_grid,dfAllTrafos_grid_PQnom, dfCombinedEvents, dfCombinedConstraints \
                        = fq.PSA_calc_flex_reqts(isAuto=False,bOutput=self.output,isBase=True, PSAFolder=PSAfolder,
                        PSArunID=PSArunID, bEvents=self.bEvents, eventsFile=self.eventFile, plout=self.plout,unplout=self.unplout,ploutfile=self.ploutFile,
                        unploutfile=self.unploutFile, 
                        bSwitch=self.bSwitch, switchFile=self.switchFile,
                        lne_thresh=self.LneThresh,trafo_thresh=self.TrafoThresh,
                        power_factor=self.powFactor,lagging=self.lagging, start_time_row=start_time_row,bDebug=pfconf.bDebug,
                        PSA_running_mode=PSA_running_mode,pf_proj_dir=Path(self.BSPmodel).parent.absolute(),start_time='',
                        dictSIAFeeder=dict(), dictSIAGen=dict(), dictSIAGroup=dict(), NeRDA_output_path=strNeRDAFolder,
                        lNamesGrid=lNamesGrid,dfDownScalingFactor_grid=pd.DataFrame(),dfAllTrafos_grid_PQnom=pd.DataFrame(),NeRDAfolder=strNeRDAFolder,
                        SIA_folder=strSIAFolder, strScenario_folder=strResultsFolder, input_folder=strInputFolder,
                        runNum=self.count, numRuns=self.NumAutoRuns)

                    if status != const.PSAok:
                        showwarning(title="WARNING", message=msg, parent=self.root)
                        self.varDisplayPSArunID.set('Process failed')
                        self.root.update_idletasks()
                    else:
                        msg = "MANUAL mode finished. \n" + msg + "\n"

                ### AUTOMATIC MODE ###
                elif self.strMode == const.strRunModeSEMIAUT:
                    strInputFolder, strResultsFolder, strNeRDAFolder, strSIAFolder = \
                        op.PSA_create_man_results_folders(PSAfolder, self.plout, self.unplout)
                    op.copy2ResultsFolder(strInputFolder, lModels)
                    status, msg, end_time_flex_reqts_aut_simpl, dictSIAFeeder, dictSIAGen, dictSIAGroup, start_time_base, \
                    NeRDA_output_path, dfDownScalingFactor_grid, dfAllTrafos_grid_PQnom, dfCombinedEvents, dfCombinedConstraints \
                        = fq.PSA_calc_flex_reqts(isAuto=True, bOutput=self.output, isBase=True,
                                                 PSAFolder=PSAfolder,
                                                 PSArunID=PSArunID, bEvents=self.bEvents, eventsFile=self.eventFile,
                                                 plout=self.plout, unplout=self.unplout, ploutfile=self.ploutFile,
                                                 unploutfile=self.unploutFile, 
                                                 bSwitch=self.bSwitch, switchFile=self.switchFile,
                                                 lne_thresh=self.LneThresh,
                                                 trafo_thresh=self.TrafoThresh,
                                                 power_factor=self.powFactor, lagging=self.lagging,
                                                 start_time_row=start_time_row, bDebug=pfconf.bDebug,
                                                 PSA_running_mode=PSA_running_mode,
                                                 pf_proj_dir=Path(self.BSPmodel).parent.absolute(), start_time='',
                                                 dictSIAFeeder=dict(), dictSIAGen=dict(), dictSIAGroup=dict(),
                                                 NeRDA_output_path=strNeRDAFolder,
                                                 lNamesGrid=lNamesGrid, dfDownScalingFactor_grid=pd.DataFrame(),
                                                 dfAllTrafos_grid_PQnom=pd.DataFrame(), NeRDAfolder=strNeRDAFolder,
                                                 SIA_folder=strSIAFolder, strScenario_folder=strResultsFolder,
                                                 input_folder=strInputFolder,
                                                 runNum=self.count, numRuns=self.NumAutoRuns)
                    if status != const.PSAok:
                        showwarning(title="WARNING", message=msg, parent=self.root)
                        self.varDisplayPSArunID.set('Process failed')
                        self.root.update_idletasks()
                    else:
                        msg = "AUTOMATIC mode finished. \n" + msg + "\n"
                    end_time_flex_reqts_aut = end_time_flex_reqts_aut_simpl
                else:
                    strNewFolder_input_data, strNewFolder_SIA, strNewFolder_NeRDA, strNewFolder_base = \
                        op.PSA_create_aut_folders_new(PSAfolder)
                    op.copy2ResultsFolder(strNewFolder_input_data, lModels)
                    status, msg_base, end_time_flex_reqts_base, dictSIAFeeder_base, dictSIAGen_base, dictSIAGroup_base, \
                        start_time_base, NeRDA_output_path_base, dfDownScalingFactor_grid_base,dfAllTrafos_grid_PQnom_base, \
                        dfCombinedEvents_base, dfCombinedConstraints_base = fq.PSA_calc_flex_reqts(isAuto=True,
                        bOutput=self.output,isBase=True, PSAFolder=PSAfolder,PSArunID=PSArunID, bEvents=self.bEvents, eventsFile=self.eventFile, 
                        plout=False, unplout=False, ploutfile=self.ploutFile, unploutfile=self.unploutFile,
                        bSwitch=self.bSwitch, switchFile=self.switchFile,
                        lne_thresh=self.LneThresh,
                        trafo_thresh=self.TrafoThresh,power_factor=self.powFactor,lagging=self.lagging,
                        start_time_row=start_time_row,bDebug=pfconf.bDebug,PSA_running_mode="B",
                        pf_proj_dir=Path(self.BSPmodel).parent.absolute(),start_time='',dictSIAFeeder=dict(),
                        dictSIAGen=dict(), dictSIAGroup=dict(),NeRDA_output_path='',
                        lNamesGrid=lNamesGrid,dfDownScalingFactor_grid=pd.DataFrame(),dfAllTrafos_grid_PQnom=pd.DataFrame(),NeRDAfolder=strNewFolder_NeRDA,
                        SIA_folder=strNewFolder_SIA,strScenario_folder=strNewFolder_base, input_folder=strNewFolder_input_data,
                        runNum=self.count, numRuns=self.NumAutoRuns)
                    if status != const.PSAok:
                        showwarning(title="WARNING", message=msg_base, parent=self.root)
                        self.varDisplayPSArunID.set('Process failed')
                        self.root.update_idletasks()
                    else:
                        msg = "Under BASE mode:" + msg_base + "\n"
                    lCombinedFlexReqts = [dfCombinedEvents_base]
                    lCombinedConstraints = [dfCombinedConstraints_base]
                    if self.plout and not self.unplout:
                        iNewNumber_pl, strNewFolder_pl = op.PSA_create_numbering_folder(PSAFolder=PSAfolder,
                                                               prev_numb=3,strFoldername="MAINT")
                        status, msg_pl, end_time_flex_reqts_pl, dictSIAFeeder, dictSIAGen, dictSIAGroup, start_time, \
                            NeRDA_output_path,dfDownScalingFactor_grid_plout,dfAllTrafos_grid_PQnom_plout, \
                            dfCombinedEvents_plout, dfCombinedConstraints_plout = fq.PSA_calc_flex_reqts(isAuto=True,
                            bOutput=self.output, isBase=False,PSAFolder=PSAfolder,PSArunID=PSArunID, bEvents=self.bEvents, eventsFile=self.eventFile,
                            plout=True, unplout=False,ploutfile=self.ploutFile,unploutfile=self.unploutFile,
                            bSwitch=self.bSwitch, switchFile=self.switchFile,
                            lne_thresh=self.LneThresh,
                            trafo_thresh=self.TrafoThresh,power_factor=self.powFactor,lagging=self.lagging,
                            start_time_row=start_time_row,bDebug=pfconf.bDebug,PSA_running_mode="M",pf_proj_dir=Path(
                            self.BSPmodel).parent.absolute(),start_time=start_time_base,dictSIAFeeder=dictSIAFeeder_base,
                            dictSIAGen=dictSIAGen_base, dictSIAGroup=dictSIAGroup_base, NeRDA_output_path = NeRDA_output_path_base,
                            lNamesGrid=lNamesGrid,dfDownScalingFactor_grid=dfDownScalingFactor_grid_base,dfAllTrafos_grid_PQnom=dfAllTrafos_grid_PQnom_base,
                            NeRDAfolder=strNewFolder_NeRDA,SIA_folder=strNewFolder_SIA,strScenario_folder=strNewFolder_pl,
                            input_folder=strNewFolder_input_data, runNum=self.count, numRuns=self.NumAutoRuns)
                        if status != const.PSAok:
                            # lCombinedFlexReqts, lCombinedConstraints = ([] for i in range(2))
                            showwarning(title="WARNING", message=msg_pl, parent=self.root)
                            self.varDisplayPSArunID.set('Process failed')
                            self.root.update_idletasks()
                        else:
                            lCombinedFlexReqts = [dfCombinedEvents_base, dfCombinedEvents_plout]
                            lCombinedConstraints = [dfCombinedConstraints_base, dfCombinedConstraints_plout]
                            end_time_flex_reqts_aut = end_time_flex_reqts_pl
                            msg = msg + "Under MAINT mode:" + msg_pl + "\n"
                    elif not self.plout and self.unplout:
                        iNewNumber_unpl, strNewFolder_unpl = op.PSA_create_numbering_folder(PSAFolder=PSAfolder,
                                                            prev_numb=3,strFoldername="CONT")
                        status, msg_unpl, end_time_flex_reqts_unpl,dictSIAFeeder, dictSIAGen, dictSIAGroup, start_time, \
                            NeRDA_output_path,dfDownScalingFactor_grid_unplout,dfAllTrafos_grid_PQnom_unplout, \
                            dfCombinedEvents_unplout, dfCombinedConstraints_unplout = fq.PSA_calc_flex_reqts(
                            isAuto=True,bOutput=self.output, isBase=False,PSAFolder=PSAfolder,PSArunID=PSArunID, bEvents=self.bEvents, eventsFile=self.eventFile,
                            plout=False,unplout=True,ploutfile=self.ploutFile,unploutfile=self.unploutFile,
                            bSwitch=self.bSwitch, switchFile=self.switchFile,
                            lne_thresh=self.LneThresh,
                            trafo_thresh=self.TrafoThresh,power_factor=self.powFactor,lagging=self.lagging,start_time_row=
                            start_time_row,bDebug=pfconf.bDebug,PSA_running_mode="C",pf_proj_dir=Path(
                            self.BSPmodel).parent.absolute(),start_time=start_time_base,dictSIAFeeder=dictSIAFeeder_base,
                            dictSIAGen=dictSIAGen_base, dictSIAGroup=dictSIAGroup_base,NeRDA_output_path=NeRDA_output_path_base,
                            lNamesGrid=lNamesGrid,dfDownScalingFactor_grid=dfDownScalingFactor_grid_base,dfAllTrafos_grid_PQnom=dfAllTrafos_grid_PQnom_base,
                            NeRDAfolder=strNewFolder_NeRDA,SIA_folder=strNewFolder_SIA,strScenario_folder=strNewFolder_unpl,
                            input_folder=strNewFolder_input_data, runNum=self.count, numRuns=self.NumAutoRuns)
                        if status != const.PSAok:
                            # lCombinedFlexReqts, lCombinedConstraints = ([] for i in range(2))
                            showwarning(title="WARNING", message=msg_unpl, parent=self.root)
                            self.varDisplayPSArunID.set('Process failed')
                            self.root.update_idletasks()
                        else:
                            lCombinedFlexReqts = [dfCombinedEvents_base, dfCombinedEvents_unplout]
                            lCombinedConstraints = [dfCombinedConstraints_base, dfCombinedConstraints_unplout]
                            end_time_flex_reqts_aut = end_time_flex_reqts_unpl
                            msg = msg + "Under CONT mode:" + msg_unpl + "\n"

                    elif self.plout and self.unplout:
                        iNewNumber_pl, strNewFolder_pl = op.PSA_create_numbering_folder(PSAFolder=PSAfolder,
                                                                prev_numb=3,strFoldername="MAINT")
                        status, msg_pl, end_time_flex_reqts_pl,dictSIAFeeder, dictSIAGen, dictSIAGroup, start_time, \
                        NeRDA_output_path, dfDownScalingFactor_grid_pl,dfAllTrafos_grid_PQnom_pl, dfCombinedEvents_pl, \
                        dfCombinedConstraints_pl = fq.PSA_calc_flex_reqts(isAuto=True,bOutput=self.output, isBase=False,
                               PSAFolder=PSAfolder,PSArunID=PSArunID, bEvents=self.bEvents, eventsFile=self.eventFile,
                               plout=True,unplout=False,ploutfile=self.ploutFile,
                               unploutfile=self.unploutFile,
                               bSwitch=self.bSwitch, switchFile=self.switchFile,
                               lne_thresh=self.LneThresh,trafo_thresh=self.TrafoThresh,
                               power_factor=self.powFactor,lagging=self.lagging,start_time_row=start_time_row,
                               bDebug=pfconf.bDebug,PSA_running_mode="M",pf_proj_dir=Path(self.BSPmodel).parent.absolute(),
                               start_time=start_time_base,dictSIAFeeder=dictSIAFeeder_base,dictSIAGen=dictSIAGen_base,
                               dictSIAGroup=dictSIAGroup_base,NeRDA_output_path=NeRDA_output_path_base,
                               lNamesGrid=lNamesGrid,dfDownScalingFactor_grid=dfDownScalingFactor_grid_base,NeRDAfolder=strNewFolder_NeRDA,
                               SIA_folder=strNewFolder_SIA,dfAllTrafos_grid_PQnom=dfAllTrafos_grid_PQnom_base,
                               strScenario_folder=strNewFolder_pl,input_folder=strNewFolder_input_data,
                               runNum=self.count, numRuns=self.NumAutoRuns)
                        if status != const.PSAok:
                            showwarning(title="WARNING", message=msg_pl, parent=self.root)
                            self.varDisplayPSArunID.set('Process failed')
                            self.root.update_idletasks()
                        else:
                            msg = msg + "Under MAINT mode:" + "\n" + msg_pl + "\n"

                        iNewNumber_unpl, strNewFolder_unpl = op.PSA_create_numbering_folder(PSAFolder=PSAfolder,
                                                            prev_numb=iNewNumber_pl,strFoldername="CONT")
                        status, msg_unpl, end_time_flex_reqts_unpl,dictSIAFeeder, dictSIAGen, dictSIAGroup, start_time, \
                        NeRDA_output_path, dfDownScalingFactor_grid_unpl,dfAllTrafos_grid_PQnom_unpl, dfCombinedEvents_unpl, \
                        dfCombinedConstraints_unpl = fq.PSA_calc_flex_reqts(isAuto=True,bOutput=self.output, isBase=False,
                             PSAFolder=PSAfolder,PSArunID=PSArunID, bEvents=self.bEvents, eventsFile=self.eventFile,
                             plout=False,unplout=True,ploutfile=self.ploutFile,
                             unploutfile=self.unploutFile,
                             bSwitch=self.bSwitch, switchFile=self.switchFile,
                             lne_thresh=self.LneThresh,trafo_thresh=self.TrafoThresh,
                             power_factor=self.powFactor,lagging=self.lagging,start_time_row=start_time_row,
                             bDebug=pfconf.bDebug,PSA_running_mode="C",pf_proj_dir=Path(self.BSPmodel).parent.absolute(),
                             start_time=start_time_base,dictSIAFeeder=dictSIAFeeder_base,dictSIAGen=dictSIAGen_base,
                             dictSIAGroup=dictSIAGroup_base, NeRDA_output_path=NeRDA_output_path_base,
                             lNamesGrid=lNamesGrid,dfDownScalingFactor_grid=dfDownScalingFactor_grid_base,NeRDAfolder=strNewFolder_NeRDA,
                             SIA_folder=strNewFolder_SIA,dfAllTrafos_grid_PQnom=dfAllTrafos_grid_PQnom_base,
                             strScenario_folder=strNewFolder_unpl,input_folder=strNewFolder_input_data,
                             runNum=self.count, numRuns=self.NumAutoRuns)
                        if status != const.PSAok:
                            showwarning(title="WARNING", message=msg_unpl, parent=self.root)
                            self.varDisplayPSArunID.set('Process failed')
                            self.root.update_idletasks()
                        else:
                            msg = msg + "Under CONT mode:" + "\n" + msg_unpl + "\n"

                        iNewNumber_pl_unpl, strNewFolder_pl_unpl = op.PSA_create_numbering_folder(PSAFolder=PSAfolder,
                                                                 prev_numb=iNewNumber_unpl,strFoldername="MAINT_CONT")
                        status, msg_pl_unpl, end_time_flex_reqts_pl_unpl,dictSIAFeeder, dictSIAGen, dictSIAGroup, \
                        start_time, NeRDA_output_path,dfDownScalingFactor_grid_pl_unpl,dfAllTrafos_grid_PQnom_pl_unpl,\
                        dfCombinedEvents_pl_unpl, dfCombinedConstraints_pl_unpl=fq.PSA_calc_flex_reqts(isAuto=True,
                            bOutput=self.output, isBase=False,PSAFolder=PSAfolder,PSArunID=PSArunID, bEvents=self.bEvents, eventsFile=self.eventFile,
                            plout=True,unplout=True,
                            ploutfile=self.ploutFile,unploutfile=self.unploutFile,
                            bSwitch=self.bSwitch, switchFile=self.switchFile,
                            lne_thresh=self.LneThresh,trafo_thresh=self.TrafoThresh,
                            power_factor=self.powFactor,lagging=self.lagging,start_time_row=start_time_row,
                            bDebug=pfconf.bDebug,PSA_running_mode="MC",pf_proj_dir=Path(self.BSPmodel).parent.absolute(),
                            start_time=start_time_base,dictSIAFeeder=dictSIAFeeder_base,dictSIAGen=dictSIAGen_base,
                            dictSIAGroup=dictSIAGroup_base,NeRDA_output_path=NeRDA_output_path_base,
                            lNamesGrid=lNamesGrid,dfDownScalingFactor_grid=dfDownScalingFactor_grid_base,NeRDAfolder=strNewFolder_NeRDA,
                            SIA_folder=strNewFolder_SIA,dfAllTrafos_grid_PQnom=dfAllTrafos_grid_PQnom_base,
                            strScenario_folder=strNewFolder_pl_unpl,input_folder=strNewFolder_input_data,
                            runNum=self.count, numRuns=self.NumAutoRuns)
                        if status != const.PSAok:
                            # lCombinedFlexReqts, lCombinedConstraints = ([] for i in range(2))
                            showwarning(title="WARNING", message=msg_pl_unpl, parent=self.root)
                            self.varDisplayPSArunID.set('Process failed')
                            self.root.update_idletasks()
                        else:
                            msg = msg + "Under MAINT_CONT mode:" + "\n" + msg_pl_unpl + "\n"
                            lCombinedFlexReqts = [dfCombinedEvents_base, dfCombinedEvents_pl, dfCombinedEvents_unpl,
                                                  dfCombinedEvents_pl_unpl]
                            lCombinedConstraints = [dfCombinedConstraints_base, dfCombinedConstraints_pl,
                                                    dfCombinedConstraints_unpl, dfCombinedConstraints_pl_unpl]
                        end_time_flex_reqts_aut = end_time_flex_reqts_pl_unpl
                    else:
                        end_time_flex_reqts_aut = end_time_flex_reqts_base
                
                if status == const.PSAok:
                    if self.strMode == const.strRunModeMAN or self.strMode == const.strRunModeSEMIAUT:
                        dfFlexReqts = dfCombinedEvents
                        dfConstraints = dfCombinedConstraints
                    else:
                        dfFlexReqts = pd.concat(lCombinedFlexReqts)
                        dfConstraints = pd.concat(lCombinedConstraints)

                    # if self.auto:
                    #     dfFlexReqts = pd.concat(lCombinedFlexReqts)
                    #     dfConstraints = pd.concat(lCombinedConstraints)
                    # else:
                    #     dfFlexReqts = dfCombinedEvents
                    #     dfConstraints = dfCombinedConstraints

                    if dfFlexReqts.empty:
                        #output NO RESULTS files to PSA and S&D (if necessary)
                        outFile = PSAfolder + "\\" + PSArunID + const.strPSANoFlexReqts + ".xlsx"
                        pd.DataFrame().to_excel(outFile,index=False)
                        if self.output:
                            outFile = self.dict_config[const.strFileDetectorFolder] + "\\" + PSArunID + const.strPSANoFlexReqts + ".xlsx"
                            pd.DataFrame().to_excel(outFile, index=False)
                        outFile = PSAfolder + "\\" + PSArunID + const.strPSANoConstraints + ".xlsx"
                        pd.DataFrame().to_excel(outFile, index=False)
                    else:
                        #output RESULTS files to PSA and S&D (if necessary)
                        dfFlexReqts.reset_index()
                        dfFlexReqts['req_id'] = [i for i in range(len(dfFlexReqts))]
                        dfFlexReqts.set_index('req_id')
                        outFile = PSAfolder + "\\" + PSArunID + const.strPSAFlexReqts + ".xlsx"
                        dfFlexReqts.to_excel(outFile, index=False)
                        if self.output:
                            outFile = self.dict_config[const.strFileDetectorFolder] + "\\" + PSArunID + const.strPSAFlexReqts + ".xlsx"
                            dfFlexReqts.to_excel(outFile, index=False)
                        dfConstraints.reset_index()
                        dfConstraints['req_id'] = [i for i in range(len(dfConstraints))]
                        dfConstraints.set_index('req_id')
                        outFile = PSAfolder + "\\" + PSArunID + const.strPSAConstraints + ".xlsx"
                        dfConstraints.to_excel(outFile, index=False)

                    if self.auto:
                        print("TOTAL TIME ELAPSED for AUTO MODE: " + str(end_time_flex_reqts_aut - start_time_flex_reqts))
                    else:
                        print("TOTAL TIME ELAPSED for MANUAL MODE: " + str(end_time_flex_reqts_manual - start_time_flex_reqts))

                print("################################################################")
                PSAstatus = status
                if status == const.PSAok:
                    if not self.auto:
                        self.enableButtons(self.fileBtn1, self.fileBtn3, self.events, self.plan, self.unpl, self.switch, self.xlBtn3,
                                    self.view_outage_data_btn, self.view_threshold_data_btn,
                                    self.switchPFLag, self.switchPFLead, 
                                    self.entryPowerFactor, self.entryLneThresh, self.entryTrafoThresh,
                                    self.switchAutomatic, self.switchAutomaticSimplified, self.switchManual, 
                                    self.calBtn, self.entryAutoRuns, self.entryNumDays)
                        showinfo(title="SUCCESS", message=msg, parent=self.root)
                        self.varDisplayPSArunID.set('Manual process finished')
                        self.root.update_idletasks()
                    elif self.auto and (self.NumAutoRuns != -1) and (self.count >= self.NumAutoRuns):
                        self.enableButtons(self.fileBtn1, self.fileBtn3, self.events, self.plan, self.unpl, self.switch, self.xlBtn3,
                                    self.view_outage_data_btn, self.view_threshold_data_btn,
                                    self.switchPFLag, self.switchPFLead, 
                                    self.entryPowerFactor, self.entryLneThresh, self.entryTrafoThresh,
                                    self.switchAutomatic, self.switchAutomaticSimplified, self.switchManual, 
                                    self.calBtn, self.entryAutoRuns, self.entryNumDays)

                        msg = "Automated process completed after " + str(self.count) + " runs"
                        pbar.updateProgressBar(100, msg)
                        showinfo(title="SUCCESS", message=msg, parent=self.root)
                        self.varDisplayPSArunID.set('Automated process finished')
                        self.root.update_idletasks()

        return PSAstatus, PSArunID

    def validUserInput(self, inVal, minVal, maxVal):
        valid = True
        fMinVal = float(minVal)
        fMaxVal = float(maxVal)
        try:
            fVal = float(inVal.get())
        except:
            valid = False
        if valid:
            valid = (fVal >= fMinVal) and (fVal <= fMaxVal)
        return valid

    def calcBtn_clicked(self):
        self.bStopProcess = False
        self.get_running_mode()
        PSAstatus = const.PSAok

        # validate numeric inputs
        self.NumAutoRuns = self.varAutoRuns.get()
        if not self.validUserInput(self.varAutoRuns, -1, 1000) or self.NumAutoRuns == 0:
            PSAstatus = const.PSAdataEntryError
            msg = "Invalid number of Automatic runs (-1 to 1000) \n"
            msg = msg + " 0 = Invalid \n"
            msg = msg + "-1 = Non-stop"

        if not self.validUserInput(self.varNumDays, 1, 11):
            PSAstatus = const.PSAdataEntryError
            msg = "Invalid number of days (1 to 11)"

        if not self.validUserInput(self.varLneThresh, 0.0, 200.0):
            PSAstatus = const.PSAdataEntryError
            msg = "Invalid line threshold (0.0 to 200.0)"

        if not self.validUserInput(self.varTrafoThresh, 0.0, 200.0):
            PSAstatus = const.PSAdataEntryError
            msg = "Invalid transformer threshold (0.0 to 200.0)"

        if not self.validUserInput(self.entryPowerFactor, 0.1, 1.0):
            PSAstatus = const.PSAdataEntryError
            msg = "Invalid power factor (0.1 to 1.0)"

        if self.fname3.get("1.0", END).strip() == const.strNoFileDisp:
            PSAstatus = const.PSAdataEntryError
            msg = "Please select an Asset Data file"

        if self.fname1.get("1.0", END).strip() == const.strNoFileDisp + " (*.pfd)":
            PSAstatus = const.PSAdataEntryError
            msg = "Please select a PF Model"

        if PSAstatus != const.PSAok:
            showwarning(title="File Entry Error", message=msg, parent=self.root)

        if (PSAstatus == const.PSAok) and askyesno(title='Info', message="Are you sure you want to start this " + self.strMode + " process?", parent=self.root):

            if self.auto:
                self.disableButtons(self.fileBtn1, self.fileBtn3, self.events, self.plan, self.unpl, self.switch, self.xlBtn3,
                                    self.view_outage_data_btn, self.view_threshold_data_btn,
                                    self.switchPFLag, self.switchPFLead, 
                                    self.entryPowerFactor, self.entryLneThresh, self.entryTrafoThresh,
                                    self.switchAutomatic, self.switchAutomaticSimplified, self.switchManual, 
                                    self.calBtn, self.entryAutoRuns, self.entryNumDays)
                self.root.update_idletasks()
                self.count = 0
                PSAstatus, PSArunID = self.repeatAuto()
            else:
                self.disableButtons(self.fileBtn1, self.fileBtn3, self.events, self.plan, self.unpl, self.switch, self.xlBtn3,
                                    self.view_outage_data_btn, self.view_threshold_data_btn,
                                    self.switchPFLag, self.switchPFLead, 
                                    self.entryPowerFactor, self.entryLneThresh, self.entryTrafoThresh,
                                    self.switchAutomatic, self.switchAutomaticSimplified, self.switchManual, 
                                    self.calBtn, self.entryAutoRuns, self.entryNumDays)
                PSAstatus, PSArunID = self.calcBtn_clicked_core()
                
            if PSAstatus == const.PSAok:
                self.varDisplayPSArunID.set(PSArunID + " processed successfully")
                self.root.update_idletasks()
                if self.auto and (self.NumAutoRuns != -1) and (self.count >= self.NumAutoRuns):
                    self.enableButtons(self.fileBtn1, self.fileBtn3, self.events, self.plan, self.unpl, self.switch, self.xlBtn3,
                                    self.view_outage_data_btn, self.view_threshold_data_btn,
                                    self.switchPFLag, self.switchPFLead, 
                                    self.entryPowerFactor, self.entryLneThresh, self.entryTrafoThresh,
                                    self.switchAutomatic, self.switchAutomaticSimplified, self.switchManual, 
                                    self.calBtn, self.entryAutoRuns, self.entryNumDays)
                    self.root.update_idletasks()
                    showinfo(title="INFO", message=PSArunID + " completed automatic processing", parent=self.root)
                elif not self.auto:
                    self.enableButtons(self.fileBtn1, self.fileBtn3, self.events, self.plan, self.unpl, self.switch, self.xlBtn3,
                                    self.view_outage_data_btn, self.view_threshold_data_btn,
                                    self.switchPFLag, self.switchPFLead, 
                                    self.entryPowerFactor, self.entryLneThresh, self.entryTrafoThresh,
                                    self.switchAutomatic, self.switchAutomaticSimplified, self.switchManual, 
                                    self.calBtn, self.entryAutoRuns, self.entryNumDays)
                    showinfo(title="INFO", message=PSArunID + " completed manual processing", parent=self.root)
            else:
                self.enableButtons(self.fileBtn1, self.fileBtn3, self.events, self.plan, self.unpl, self.switch, self.xlBtn3,
                                self.view_outage_data_btn, self.view_threshold_data_btn,
                                self.switchPFLag, self.switchPFLead, 
                                self.entryPowerFactor, self.entryLneThresh, self.entryTrafoThresh,
                                self.switchAutomatic, self.switchAutomaticSimplified, self.switchManual, 
                                self.calBtn, self.entryAutoRuns, self.entryNumDays)                
                showinfo(title="INFO", message="Current run terminated", parent=self.root)

    def folderBtn_clicked(self, fLabel, fpat, fBtn, bInput):
        if bInput:
            file_loc = self.dict_config[const.strWorkingFolder] + const.strDataInput
        else:
            file_loc = self.dict_config[const.strWorkingFolder] + const.strDataResults

        fldr = fd.askdirectory(title="Select PSA folder", initialdir=file_loc, parent=self.root)
        fldr_extnd = fldr.replace(self.dict_config[const.strWorkingFolder], '')
        if fldr != '':
            fLabel['state'] = 'normal'
            fLabel.delete("1.0", "end")
            fLabel.insert("1.0", fldr_extnd)
            fLabel['state'] = 'disabled'
        if bool(fBtn):
            if (fLabel.get("1.0", "end").strip() != const.strNoFolderDisp) and (
                    fLabel.get("1.0", "end").strip() != const.strNoFolderDisp):
                fBtn['state'] = 'normal'
            else:
                fBtn['state'] = 'disabled'


    def fileBtn_clicked(self, fLabel, fpat, fBtn, bInput):
        if ".pfd" in str(fLabel.get("1.0", END)):
            ftypes = [('PowerFactory Files', '.pfd')]
        else:
            ftypes = [('Excel files', '.xlsx')]
        
        if bInput:
            file_loc = self.dict_config[const.strWorkingFolder] + const.strDataInput
        else:
            file_loc = self.dict_config[const.strWorkingFolder] + const.strDataResults
        
        filename = fd.askopenfilename(title='Select file', initialdir=file_loc, filetypes=ftypes, initialfile=fpat, parent=self.root)
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

    def isChecked(self, var, fileBtn):
        if var.get():
            fileBtn['state'] = NORMAL
        else:
            fileBtn['state'] = DISABLED

    def threshold_graph_btn_clicked(self):
        status = const.PSAok
        asset_data = f'{self.dict_config[const.strWorkingFolder]}{self.fname3.get("1.0","end").strip()}'
        if not os.path.isfile(asset_data):
            raise FileExistsError(f'File {asset_data} does not exist')
        tempfile = pa.plot_asset_data('', asset_data, relevant=False, save=True) # Change relevant to True to truncate results to where Events == True
        if os.path.isfile(tempfile):
            os.system('"' + tempfile + '"')
        else:
            status = const.PSAfileExistError
            msg = "Def threshold_graph_btn_clicked: File doesn't exist: " + tempfile
        if status != const.PSAok:
            showwarning(title="WARNING", message=msg, parent=self.root)


    def outage_graph_btn_clicked(self):
        status = const.PSAok
        maint, cont, event, switch = '','','',''
        if self.planvar.get():
            maint = f'{self.dict_config[const.strWorkingFolder]}{self.fname4.get("1.0", "end").strip()}'
        if self.unplvar.get():
            cont = f'{self.dict_config[const.strWorkingFolder]}{self.fname5.get("1.0", "end").strip()}'
        if self.eventsvar.get():
            event = f'{self.dict_config[const.strWorkingFolder]}{self.fname9.get("1.0", "end").strip()}'
        if self.switchvar.get(): # Uncomment when switches become available
            switch = f'{self.dict_config[const.strWorkingFolder]}{self.fname6.get("1.0", "end").strip()}'
        try:
            tempfile = pa.plot_gantt_outages_by_file(maint_file=maint, cont_file=cont, event_file=event, switch_file=switch, save=True)
            if os.path.isfile(tempfile):
                os.system('"' + tempfile + '"')
            else:
                status = const.PSAfileExistError
                msg = "Def Low_level_overview: File doesn't exist: " + tempfile
        except BaseException as err:
            status = const.PSAfileExistError
            msg = "Error viewing data"
            print(err)
        finally:
            if status != const.PSAok:
                showwarning(title="WARNING", message=msg, parent=self.root)

    def isFile_plan(self, fname4, xlBtn4):
        if fname4.get() != const.strNoFileDisp:
            xlBtn4['state'] = NORMAL
        else:
            xlBtn4['state'] = DISABLED

    def sel(self, var):
        selection = str(var.get())



    def calcResCons(self):
        sFname = self.fname1.get("1.0", END).strip()
        self.SNDfile = self.dict_config[const.strWorkingFolder] + sFname
        if sFname != const.strNoFileDisp:
            status, msg = cr.PSA_constraint_resolution(self.var2SND.get(), False, self.SNDfile)
            if status != const.PSAok:
                showwarning(title="WARNING", message=msg, parent=self.root)
                self.root.update_idletasks()
            else:
                msg = msg + " - Process successfully completed"
                showinfo(title="SUCCESS", message=msg, parent=self.root)
                self.root.update_idletasks()            
        return

    def __constraint_resolution(self):
        # ROOT WINDOW #
        self.root = tk.Toplevel(master=self, width=850, height=680)
        self.dict_config = ut.PSA_SND_read_config(const.strConfigFile)
        self.pic_excel_logo = PhotoImage(file=self.dict_config[const.strWorkingFolder] + const.strExcelLogo)
        self.pic_excel_logo = self.pic_excel_logo.subsample(120)
        self.root.title("Constraint Resolution")
        self.fname1 = Text(self.root, width=66, height=1)
        self.fname1.insert("1.0", const.strNoFileDisp)
        self.fname1['state'] = 'disabled'
        self.fname1.grid(column=1, row=0, padx=10, pady=5, sticky=tk.W)
        self.fileBtn1 = ttk.Button(self.root, text="Select S&D Candidates", command=lambda:
                        self.fileBtn_clicked(self.fname1, '*SND_CAND*', self.xlBtn1, False), state=NORMAL)
        self.fileBtn1.grid(column=0, row=0, padx=10, pady=5, sticky=tk.W)
        self.xlBtn1 = ttk.Button(self.root, text="Edit Data",
                                 command=lambda: ut.open_xl(self.fname1, self.dict_config[const.strWorkingFolder]),
                                 image=self.pic_excel_logo, compound=LEFT, state=DISABLED)
        self.xlBtn1.grid(column=2, row=0, padx=10, pady=0)

        self.var2SND = BooleanVar()
        self.var2SND.set(True)

        self.switch2SND = ttk.Checkbutton(self.root, text="Output results to S&D Tool?", variable=
                                            self.var2SND, onvalue=True, offvalue=False,state=NORMAL)

        self.switch2SND.grid(column=0, row=2, sticky=tk.NW, padx=10, pady=10)

        self.SFBtn = tk.Button(self.root, text="Calculate Residual Constraints", command=lambda:
                                self.calcResCons(), state=NORMAL, bg="green", fg="white")
        self.SFBtn.grid(column=1, row=2, padx=200, pady=5, sticky=tk.W, columnspan=2)

        self.fname2 = Text(self.root, width=66, height=1)
        self.fname2.insert("1.0", const.strNoFileDisp)
        self.fname2['state'] = 'disabled'
        self.fname2.grid(column=1, row=4, padx=10, pady=5, sticky=tk.W)
        
        self.fileBtn2 = ttk.Button(self.root, text="View Results", 
                                command=lambda:self.fileBtn_clicked(self.fname2, '*PSA_CAND*', self.xlBtn2, False), state=NORMAL)
        self.fileBtn2.grid(column=0, row=4, padx=10, pady=5, sticky=tk.W)
        self.xlBtn2 = ttk.Button(self.root, text="Edit Data",
                                command=lambda: ut.open_xl(self.fname2, self.dict_config[const.strWorkingFolder]),
                                image=self.pic_excel_logo, compound=LEFT, state=DISABLED)
        self.xlBtn2.grid(column=2, row=4, padx=10, pady=0)

        self.exitBtn = tk.Button(self.root, text="Exit", command=self.exitBtn_clicked, bg="black", fg="white")
        self.exitBtn.grid(column=1, row=5, columnspan=2, padx=200, pady=5, sticky=W)

        # Main Loop #
        self.root.transient(self)
        self.root.protocol("WM_DELETE_WINDOW", self.exitBtn_clicked)
        

#Sensitivity factors menu
    def __calc_SF(self):
        # ROOT WINDOW #
        self.root = tk.Toplevel(master=self, width=850, height=680)
        self.dict_config = ut.PSA_SND_read_config(const.strConfigFile)
        self.pic_excel_logo = PhotoImage(file=self.dict_config[const.strWorkingFolder] + const.strExcelLogo)
        self.pic_excel_logo = self.pic_excel_logo.subsample(120)
        self.root.title("Calculate Sensitivity Factors")
        self.fname1 = Text(self.root, width=66, height=1)
        self.fname1.insert("1.0", const.strNoFileDisp)
        self.fname1['state'] = 'disabled'
        self.fname1.grid(column=1, row=0, padx=10, pady=5, sticky=tk.W)
        self.fileBtn1 = ttk.Button(self.root, text="Select S&D Responses/Contracts", command=lambda:
                        self.fileBtn_clicked(self.fname1, '*SND*', self.xlBtn1, False), state=NORMAL)
        self.fileBtn1.grid(column=0, row=0, padx=10, pady=5, sticky=tk.W)
        self.xlBtn1 = ttk.Button(self.root, text="Edit Data",
                                 command=lambda: ut.open_xl(self.fname1, self.dict_config[const.strWorkingFolder]),
                                 image=self.pic_excel_logo, compound=LEFT, state=DISABLED)
        self.xlBtn1.grid(column=2, row=0, padx=10, pady=0)

        self.var2SND = BooleanVar()
        self.var2SND.set(True)

        self.switch2SND = ttk.Checkbutton(self.root, text="Output results to S&D Tool?", variable=
                                            self.var2SND, onvalue=True, offvalue=False,state=NORMAL)

        self.switch2SND.grid(column=0, row=2, sticky=tk.NW, padx=10, pady=10)

        self.SFBtn = tk.Button(self.root, text="Calculate Sensitivity Factors", command=lambda:
                                self.cal_SF_core(), state=NORMAL, bg="green", fg="white")
        self.SFBtn.grid(column=1, row=2, padx=200, pady=5, sticky=tk.W, columnspan=2)

        self.fname2 = Text(self.root, width=66, height=1)
        self.fname2.insert("1.0", const.strNoFileDisp)
        self.fname2['state'] = 'disabled'
        self.fname2.grid(column=1, row=4, padx=10, pady=5, sticky=tk.W)
        
        self.fileBtn2 = ttk.Button(self.root, text="View Sensitivity Factors Results", command=lambda:
        self.fileBtn_clicked(self.fname2, '*PSA_SF*', self.xlBtn2, False), state=NORMAL)
        self.fileBtn2.grid(column=0, row=4, padx=10, pady=5, sticky=tk.W)
        self.xlBtn2 = ttk.Button(self.root, text="Edit Data",
                                 command=lambda: ut.open_xl(self.fname2, self.dict_config[const.strWorkingFolder]),
                                 image=self.pic_excel_logo, compound=LEFT, state=DISABLED)
        self.xlBtn2.grid(column=2, row=4, padx=10, pady=0)

        self.exitBtn = tk.Button(self.root, text="Exit", command=self.exitBtn_clicked, bg="black", fg="white")
        self.exitBtn.grid(column=1, row=5, columnspan=2, padx=200, pady=5, sticky=W)

        # Main Loop #
        self.root.transient(self)
        self.root.protocol("WM_DELETE_WINDOW", self.exitBtn_clicked)


    def cal_SF_core(self):
        self.SND_response_file = self.dict_config[const.strWorkingFolder] + self.fname1.get("1.0", END).strip()
        status, msg = pfsf.PSA_PF_workflowSF(self.var2SND.get(), False, self.SND_response_file,
                                             verbose=pfconf.bDebug)
        if status != const.PSAok:
            showwarning(title="WARNING", message=msg, parent=self.root)
            self.root.update_idletasks()
        else:
            showinfo(title="SUCCESS", message="Sensitivity factors have been successfully calculated", parent=self.root)
            self.root.update_idletasks()
        return

    def __create_Constraints(self):
        # ROOT WINDOW #
        ws = 880
        hs = 650
        self.root = tk.Toplevel(master=self, width=ws, height=hs)
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw/2) - (ws/2)
        y = (sh/2) - (hs/2)
        wstr = str(ws) + "x" + str(hs) + "+" + str(int(x)) + "+" + str(int(y))
        self.root.geometry(wstr)
        self.root.resizable(0,0)
        #AM#
        #self.root.attributes("-topmost", True)
        self.root.title("Calculate Flexibility Requirements")
        self.bStopProcess = False
        self.dict_config = ut.PSA_SND_read_config(const.strConfigFile)
        self.varDisplayPSArunID = StringVar()
        self.varDisplayPSArunID.set('Process not started')
        # LABEL FRAME 1 #
        self.LF1 = ttk.LabelFrame(self.root)
        self.LF1.grid(column=0, row=2, padx=0, pady=5, sticky=tk.N)
        # LABEL FRAME 2 #
        self.LF2 = ttk.LabelFrame(self.root)
        self.LF2.grid(column=0, row=4, padx=0, pady=5, sticky=tk.N)
        displayPSArunID = Label(self.LF2, textvariable=self.varDisplayPSArunID).grid(column=2, row=1, padx=100, pady=0)
        # Label Frame 1 #
        self.pic_excel_logo = PhotoImage(file=self.dict_config[const.strWorkingFolder] + const.strExcelLogo)
        self.pic_excel_logo = self.pic_excel_logo.subsample(120)
        self.lf1 = ttk.LabelFrame(self.root, text="MODELS")
        self.lf1.grid(column=0, row=0, padx=30, pady=10, sticky=tk.NW)
        self.fileBtn1 = ttk.Button(self.lf1, text="Select PF Model", command=lambda: self.fileBtn_clicked(self.fname1, '', '', True),
                                   state=NORMAL)
        self.fname1 = Text(self.lf1, width=66, height=1)
        #Removed need for PSA_SND_BSP_MODEL in config.txt
        pf_model_dir = const.strNoFileDisp + " (*.pfd)"
        self.fname1.insert("1.0", pf_model_dir)
        self.fname1['state'] = 'disabled'

        self.labelPrimary = Label(self.lf1, text='Select Primary: ').grid(row=3, column=0, sticky=W, padx=10,
                                                                                 pady=5)
        self.valsvar0 = tk.StringVar()
        self.valBSPs = ("All selected", "BERI", "COLO", "KENN", "ROSH", "WALL")
        self.cboxBSPs = ttk.Combobox(self.lf1, textvariable=self.valsvar0)
        self.cboxBSPs['values'] = self.valBSPs
        self.cboxBSPs.current(0)
        self.cboxBSPs['state'] = 'readonly'
        self.cboxBSPs.bind('<<ComboboxSelected>>')
        self.valsvar0.set("All selected")
        self.cboxBSPs['state'] = 'disabled'
        self.cboxBSPs.grid(column=2, row=3, padx=10, pady=5, sticky=tk.NW, columnspan=2)

        self.fileBtn3 = ttk.Button(self.lf1, text="Select Asset Data",
                                   command=lambda: self.fileBtn_clicked(self.fname3, '*ASSET*', self.xlBtn3, True))
        self.fname3 = Text(self.lf1, width=66, height=1)
        self.fname3.insert("1.0", const.strNoFileDisp)
        self.fname3['state'] = 'disabled'
        self.xlBtn3 = ttk.Button(self.lf1, text="Edit Data",
                                 command=lambda: ut.open_xl(self.fname3, self.dict_config[const.strWorkingFolder]),
                                 image=self.pic_excel_logo, compound=LEFT, state=DISABLED)




        self.fileBtn1.grid(column=0, row=2, padx=10, pady=5, sticky=tk.W)
        #self.fileBtn2.grid(column=0, row=3, padx=10, pady=5, sticky=tk.W)
        self.fileBtn3.grid(column=0, row=4, padx=10, pady=5, sticky=tk.W)
        self.xlBtn3.grid(column=4, row=4, padx=10, pady=5, sticky=tk.W)
        self.fname1.grid(column=3, row=2, padx=10, pady=5, sticky=tk.W)
        #self.fname2.grid(column=3, row=3, padx=10, pady=5, sticky=tk.W)
        self.fname3.grid(column=3, row=4, padx=10, pady=5, sticky=tk.W)
                
        # Label Frame 3 #
        self.lf3 = ttk.LabelFrame(self.root, text="SCENARIOS")
        self.lf3.grid(column=0, row=1, padx=30, pady=5, sticky=tk.NW, columnspan=2)
        self.basevar = BooleanVar(value=True)
        self.eventsvar = BooleanVar(value=False)
        self.planvar = BooleanVar(value=False)
        self.unplvar = BooleanVar(value=False)
        self.switchvar = BooleanVar(value=False)

        self.fileBtn9 = ttk.Button(self.lf3, text="Select Event Data",
                                   command=lambda: self.fileBtn_clicked(self.fname9, '*EVENT*', self.xlBtn9, True), state=DISABLED)
        self.fname9 = Text(self.lf3, width=50, height=1)
        self.fname9.insert("1.0", const.strNoFileDisp)
        self.fname9['state'] = 'disabled'
        self.xlBtn9 = ttk.Button(self.lf3, text="Edit Data",
                                 command=lambda: ut.open_xl(self.fname9, self.dict_config[const.strWorkingFolder]),
                                 image=self.pic_excel_logo, compound=LEFT, state=DISABLED)        
        self.fileBtn9.grid(column=2, row=3, padx=10, pady=5, sticky=tk.W)
        self.fname9.grid(column=3, row=3, padx=10, pady=5, sticky=tk.W)
        self.xlBtn9.grid(column=4, row=3, padx=10, pady=5, sticky=tk.W)

        self.fileBtn4 = ttk.Button(self.lf3, text="Select Maintenance Data", command=lambda: self.fileBtn_clicked(
            self.fname4, '*MAINT*', self.xlBtn4, True), state=DISABLED)
        self.fname4 = Text(self.lf3, width=50, height=1)
        self.fname4.insert("1.0", const.strNoFileDisp)
        self.fname4['state'] = 'disabled'
        self.xlBtn4 = ttk.Button(self.lf3, text="Edit Data",
                                 command=lambda: ut.open_xl(self.fname4, self.dict_config[const.strWorkingFolder]),
                                 image=self.pic_excel_logo, compound=LEFT, state=DISABLED)
        self.xlBtn4.grid(column=4, row=4, padx=10, pady=0)

        self.fileBtn5 = ttk.Button(self.lf3, text="Select Contingency Data", command=lambda: self.fileBtn_clicked(
            self.fname5, '*CONT*', self.xlBtn5, True), state=DISABLED)

        self.fname5 = Text(self.lf3, width=50, height=1)
        self.fname5.insert("1.0", const.strNoFileDisp)
        self.fname5['state'] = 'disabled'
        self.xlBtn5 = ttk.Button(self.lf3, text="Edit Data",
                                 command=lambda: ut.open_xl(self.fname5, self.dict_config[const.strWorkingFolder]),
                                 image=self.pic_excel_logo, compound=LEFT, state=DISABLED)
        self.xlBtn5.grid(column=4, row=5, padx=10, pady=0)

        self.fileBtn4.grid(column=2, row=4, padx=10, pady=5, sticky=tk.W)
        self.fname4.grid(column=3, row=4, padx=10, pady=5, sticky=tk.W)
        self.fileBtn5.grid(column=2, row=5, padx=10, pady=5, sticky=tk.W)
        self.fname5.grid(column=3, row=5, padx=10, pady=5, sticky=tk.W)

        self.base = ttk.Checkbutton(self.lf3, text="Base model", variable=self.basevar, onvalue=True, offvalue=False,
                                    state=DISABLED)
        self.events = ttk.Checkbutton(self.lf3, text="Events", variable=self.eventsvar, onvalue=True, offvalue=False,
                                    command=lambda: self.isChecked(self.eventsvar, self.fileBtn9), state=NORMAL)
        self.plan = ttk.Checkbutton(self.lf3, text="Maintenance", variable=self.planvar, onvalue=True, offvalue=False,
                                    command=lambda: self.isChecked(self.planvar, self.fileBtn4), state=NORMAL)
        self.unpl = ttk.Checkbutton(self.lf3, text="Contingency", variable=self.unplvar, onvalue=True, offvalue=False,
                                    command=lambda: self.isChecked(self.unplvar, self.fileBtn5), state=NORMAL)
        
        if pfconf.bSwitching:
            self.switch = ttk.Checkbutton(self.lf3, text="Switching", variable=self.switchvar, onvalue=True, offvalue=False,
                                    command=lambda: self.isChecked(self.switchvar, self.fileBtn6), state=NORMAL)
        else:
            self.switch = ttk.Checkbutton(self.lf3, text="Switching", variable=self.switchvar, onvalue=True, offvalue=False,
                                    command=lambda: self.isChecked(self.switchvar, self.fileBtn6), state=DISABLED)

        self.base.grid(column=0, row=2, sticky=tk.W, padx=10)
        self.events.grid(column=0, row=3, sticky=tk.W, padx=10)
        self.plan.grid(column=0, row=4, sticky=tk.W, padx=10)
        self.unpl.grid(column=0, row=5, sticky=tk.W, padx=10)
        self.switch.grid(column=0, row=6, sticky=tk.W, padx=10)

        #SWITCH DATA
        self.fileBtn6 = ttk.Button(self.lf3, text="Select Switch Data",
                                   command=lambda: self.fileBtn_clicked(self.fname6, '*SWITCH*', self.xlBtn6, True), state=DISABLED)
        self.fname6 = Text(self.lf3, width=50, height=1)
        if pfconf.bSwitching:
            self.fname6.insert("1.0", const.strNoFileDisp)
        else:
            self.fname6.insert("1.0", "Not yet implemented")
        self.fname6['state'] = 'disabled'
        self.xlBtn6 = ttk.Button(self.lf3, text="Edit Data",
                                 command=lambda: ut.open_xl(self.fname6, self.dict_config[const.strWorkingFolder]),
                                 image=self.pic_excel_logo, compound=LEFT, state=DISABLED)        
        self.fileBtn6.grid(column=2, row=6, padx=10, pady=5, sticky=tk.W)
        self.fname6.grid(column=3, row=6, padx=10, pady=5, sticky=tk.W)
        self.xlBtn6.grid(column=4, row=6, padx=10, pady=5, sticky=tk.W)

        # Activity and Threshold graphs
        self.view_outage_data_btn = tk.Button(self.lf3, text='View Activity', command=self.outage_graph_btn_clicked, 
                                              state=NORMAL, bg='white',fg='green')
        self.view_outage_data_btn.grid(column=0, row=7, padx=10, pady=10, sticky=tk.W)

        self.view_threshold_data_btn = tk.Button(self.lf3, text=' View Thresholds', command=self.threshold_graph_btn_clicked, 
                                                 state=NORMAL, bg='white',fg='blue')
        self.view_threshold_data_btn.grid(column=2,row=7, padx=10, pady=10, sticky=tk.W)

        # Label Frame 4 #
        self.LF3 = ttk.LabelFrame(self.root, text="")
        self.LF3.grid(column=0, row=3, padx=10, pady=0, columnspan=2)

        self.lf8 = ttk.LabelFrame(self.LF3, text="POWER FACTOR")
        self.lf8.grid(column=0, row=0, padx=10, pady=5, columnspan=2,sticky=tk.NW)

        self.lf9 = ttk.LabelFrame(self.LF3, text="THRESHOLDS")
        self.lf9.grid(column=2, row=0, padx=10, pady=5, columnspan=2, sticky=tk.NW)

        r=0

        self.labelPowerFactor = Label(self.lf8, text='Power factor: ').grid(row=r, column=0, sticky=W, padx=10, pady=5)
        self.varPowerFactor = DoubleVar(self.lf8, value=0.95)
        self.entryPowerFactor = Entry(self.lf8, textvariable=self.varPowerFactor, state=NORMAL, width=5)
        self.entryPowerFactor.grid(column=1, row=r, sticky=tk.NW, padx=10, pady=5)
        r=r+1

        self.varPF = BooleanVar()
        self.varPF.set(True)

        self.switchPFLag = Radiobutton(self.lf8, text='Lagging', variable=self.varPF, value=True,
                                       command=lambda: self.sel(self.varPF))
        self.switchPFLag.grid(column=0, row=r, sticky=tk.NW, padx=10, pady=0)
        self.switchPFLead = Radiobutton(self.lf8, text='Leading', variable=self.varPF, value=False,
                                        command=lambda: self.sel(self.varPF))
        self.switchPFLead.grid(column=1, row=r, sticky=tk.NW, padx=10, pady=0)
        r=r+1

        self.labelLneThresh = Label(self.lf9, text='Default line loading threshold (%): ')
        self.labelLneThresh.grid(column=0, row=r, sticky=W, padx=10,pady=5)                                                                                    
        self.varLneThresh = DoubleVar(self.lf9, value=100)
        self.entryLneThresh = Entry(self.lf9, textvariable=self.varLneThresh, state=NORMAL, width=5)
        self.entryLneThresh.grid(column=1, row=r, sticky=tk.NW, padx=10, pady=5)
        r=r+1

        self.labelTrafoThresh = Label(self.lf9, text='Default transformer loading threshold (%): ')
        self.labelTrafoThresh.grid(column=0, row=r, sticky=W,padx=10,pady=5)                                
        self.varTrafoThresh = DoubleVar(self.lf9, value=100)
        self.entryTrafoThresh = Entry(self.lf9, textvariable=self.varTrafoThresh, state=NORMAL, width=5)
        self.entryTrafoThresh.grid(column=1, row=r, sticky=tk.NW, padx=10, pady=5)
        r=r+1

        self.lf7 = ttk.LabelFrame(self.LF3, text="ADVANCED SETTINGS")
        self.lf7.grid(column=4, row=0, columnspan=1, padx=10, pady=5, sticky=tk.NW)
        self.labelAutoRuns = Label(self.lf7, text='Number of Automatic runs: ')
        self.labelAutoRuns.grid(column=0, row=0, sticky=W, padx=10,pady=5)                                                                                    
        self.varAutoRuns = IntVar(self.lf7, value=pfconf.autoNumRuns)
        self.entryAutoRuns = Entry(self.lf7, textvariable=self.varAutoRuns, state=NORMAL, width=5)
        self.entryAutoRuns.grid(column=1, row=0, sticky=tk.NW, padx=10, pady=5)
        self.labelNumDays = Label(self.lf7, text='Number of days: ')
        self.labelNumDays.grid(column=0, row=1, sticky=W, padx=10,pady=5)                                                                                    
        self.varNumDays = IntVar(self.lf7, value=pfconf.number_of_days)
        self.entryNumDays = Entry(self.lf7, textvariable=self.varNumDays, state=NORMAL, width=5)
        self.entryNumDays.grid(column=1, row=1, sticky=tk.NW, padx=10, pady=5)



        # Label Frame  5 #
        self.lf5 = ttk.LabelFrame(self.LF2, text="PROCESSING MODE")
        self.lf5.grid(column=0, row=0, columnspan=2, padx=10, pady=0, sticky=tk.NW)

        self.v = IntVar()
        self.switchAutomatic = tk.Radiobutton(self.lf5,text="Automatic",padx=20,variable=self.v,value=1,
                                              command=lambda: self.sel(self.v))
        self.switchAutomatic.grid(column=0, row=0, sticky=tk.NW, padx=10, pady=0)
        self.switchAutomaticSimplified = tk.Radiobutton(self.lf5, text="Automatic (Simplified)", padx=20,variable=self.v,
                                                        value=2,command=lambda: self.sel(self.v))
        self.switchAutomaticSimplified.grid(column=0, row=1, sticky=tk.NW, padx=10, pady=0)
        self.switchManual = tk.Radiobutton(self.lf5, text="Manual", padx=20,variable=self.v, value=3,
                                           command=lambda: self.sel(self.v))
        self.v.set(3)  # initializing the choice, i.e. Manual
        self.switchManual.grid(column=0, row=2, sticky=tk.NW, padx=10, pady=0)


        self.lf6 = ttk.LabelFrame(self.LF2, text="OUTPUTS")
        self.lf6.grid(column=2, row=0, columnspan=1, padx=10, pady=0, sticky=tk.NW)

        self.calBtn = tk.Button(self.lf6, text="Calculate Flexibility Requirements", command=self.calcBtn_clicked,
                                bg="green", fg="white")
        self.calBtn.grid(column=0, row=0, columnspan=1, padx=10, pady=10)

        # Buttons #
        self.outputvar = BooleanVar(value=True)
        self.output = ttk.Checkbutton(self.lf6, text="Output results to S&D Tool", variable=self.outputvar, onvalue=True,
                                      offvalue=False,state=NORMAL)
        self.output.grid(column=0, row=1, sticky=tk.W, padx=10, pady=0)

        self.stopBtn = tk.Button(self.LF2, text="Stop", command=self.stopBtn_clicked, bg="red", fg="white")
        self.stopBtn.grid(column=3, row=0, columnspan=1, padx=10, pady=0, sticky=W)

        self.exitBtn = tk.Button(self.LF2, text="Exit", command=self.exitBtn_clicked, bg="black", fg="white")
        self.exitBtn.grid(column=4, row=0, columnspan=1, padx=10, pady=0, sticky=W)

        # Main Loop #
        self.root.transient(self)
        self.root.protocol("WM_DELETE_WINDOW", self.exitBtn_clicked)

    def processData1(self, outFile, calcHist, otmBox, infile, mFile, top_folder, var2SND1, var2SND2):
        status = const.PSAok
        msg = "processData1"
        psaFile1 = infile.get("1.0", "end").strip()
        psaFile = str(psaFile1)
        if psaFile != const.strNoFileDisp:
            psaFile = top_folder + psaFile
        mapFile = str(mFile.get("1.0", "end").strip())
        if mapFile != const.strNoFileDisp:
            mapFile = top_folder + mapFile
        outType = str(outFile.get())
        cHist = str(calcHist.get())
        otm = str(otmBox.get())
        trunc = bool(var2SND2.get())
        out_to_snd = bool(var2SND1.get())
        accept = bool(self.valsvar3.get())

        status,msg,snd_file_location = SNDfile.create_snd(psaFile, mapFile, otm, cHist, file_type=outType, out_to_snd=out_to_snd, truncate=trunc, accepted=accept)
        if status != const.PSAok:
            showinfo(title="ERROR", message=msg, parent=self.root)
        else:
            showinfo(title="INFO", message=f'Saved to: {snd_file_location}', parent=self.root)
        return status, msg

    def processData2(self, outFile, calcHist, infile, top_folder):
        status = const.PSAok
        msg = "processData2"
        psaFile = str(infile.get("1.0", "end").strip())
        if psaFile != const.strNoFileDisp:
            psaFile = top_folder + psaFile
        outType = str(outFile.get())
        cHist = str(calcHist.get())
        accept = bool(self.valsvar3.get())
        try:
            status,msg,snd_file_location = SNDfile.create_candidates(psaFile, cHist, outType, accept)
        except BaseException as err:
            status = const.PSAdataEntryError
            msg = err
        if status != const.PSAok:
            showinfo(title="ERROR", message=msg, parent=self.root)
        else:
            showinfo(title="INFO", message=f'Saved to: {snd_file_location}', parent=self.root)
        return status, msg

    def __createSNDfiles(self):
        # ROOT WINDOW #
        self.root = tk.Toplevel(master=self, width=850, height=680)
        self.dict_config = ut.PSA_SND_read_config(const.strConfigFile)
        self.pic_excel_logo = PhotoImage(file=self.dict_config[const.strWorkingFolder] + const.strExcelLogo)
        self.pic_excel_logo = self.pic_excel_logo.subsample(120)
        self.root.title("Create S&D Output Files")
        
        #SELECT FLEX-REQTS FILE
        self.fname1 = Text(self.root, width=66, height=1)
        self.fname1.insert("1.0", const.strNoFileDisp)
        self.fname1['state'] = 'disabled'
        self.fname1.grid(column=1, row=0, padx=10, pady=5, sticky=tk.W)
        self.fileBtn1 = ttk.Button(self.root, text="Select PSA Flex Reqts file", command=lambda:
                        self.fileBtn_clicked(self.fname1, '*PSA_FLEX-*', self.xlBtn1, False), state=NORMAL)
        self.fileBtn1.grid(column=0, row=0, padx=10, pady=5, sticky=tk.W)
        topFldr = self.dict_config[const.strWorkingFolder]        
        self.xlBtn1 = ttk.Button(self.root, text="Edit Data",
                                 command=lambda: ut.open_xl(self.fname1, self.dict_config[const.strWorkingFolder]),
                                 image=self.pic_excel_logo, compound=LEFT, state=DISABLED)
        self.xlBtn1.grid(column=2, row=0, padx=10, pady=0)

        #SELECT ASSET MAPPING FILE
        self.fname2 = Text(self.root, width=66, height=1)
        self.fname2.insert("1.0", const.strNoFileDisp)
        self.fname2['state'] = 'disabled'
        self.fname2.grid(column=1, row=1, padx=10, pady=5, sticky=tk.W)
        self.fileBtn2 = ttk.Button(self.root, text="Select asset mapping file", command=lambda:
                        self.fileBtn_clicked(self.fname2, '*mapping*', self.xlBtn2, True), state=NORMAL)
        self.fileBtn2.grid(column=0, row=1, padx=10, pady=5, sticky=tk.W)
        topFldr = self.dict_config[const.strWorkingFolder]        
        self.xlBtn2 = ttk.Button(self.root, text="Edit Data",
                                 command=lambda: ut.open_xl(self.fname2, self.dict_config[const.strWorkingFolder]),
                                 image=self.pic_excel_logo, compound=LEFT, state=DISABLED)
        self.xlBtn2.grid(column=2, row=1, padx=10, pady=0)

        # Drop down boxes
        r=3
        # RESPONSES OR CONTRACTS
        self.valsvar1 = tk.StringVar()
        self.val1 = ("RESPONSES", "CONTRACTS")
        self.cbox1 = ttk.Combobox(self.root, textvariable=self.valsvar1)
        self.cbox1['values'] = self.val1
        self.cbox1.current(0)
        self.cbox1['state'] = 'readonly'
        self.cbox1.bind('<<ComboboxSelected>>')
        self.valsvar1.set("RESPONSES")
        self.cbox1.grid(column=1, row=r, padx=0, pady=5, sticky=tk.NW, columnspan=1)
        self.label1 = Label(self.root, text="Output type:")
        self.label1.grid(column=0, row=r, padx=0, pady=5, sticky=tk.NE, columnspan=1)
        # CALCULATE HISTORICAL TRUE/FALSE
        self.valsvar2 = tk.StringVar()
        self.val2 = ("TRUE", "FALSE")
        self.cbox2 = ttk.Combobox(self.root, textvariable=self.valsvar2)
        self.cbox2['values'] = self.val2
        self.cbox2.current(0)
        self.cbox2['state'] = 'readonly'
        self.cbox2.bind('<<ComboboxSelected>>')
        self.valsvar2.set("TRUE")
        self.cbox2.grid(column=1, row=r+1, padx=0, pady=5, sticky=tk.NW, columnspan=1)
        self.label2 = Label(self.root, text="Calc historical:")
        self.label2.grid(column=0, row=r+1, padx=0, pady=5, sticky=tk.NE, columnspan=1)
        # ACCEPTED TRUE/FALSE
        self.valsvar3 = tk.StringVar(value="FALSE")
        self.val3 = ("TRUE", "FALSE")
        self.cbox3 = ttk.Combobox(self.root, textvariable=self.valsvar3, values=self.val3, state='readonly')
        self.cbox3.bind('<<ComboboxSelected>>')
        self.cbox3.grid(column=1, row=r+2, padx=0, pady=5, sticky=tk.NW, columnspan=1)
        self.label3 = Label(self.root, text="Accepted:")
        self.label3.grid(column=0, row=r+2, padx=0, pady=5, sticky=tk.NE, columnspan=1)
        # RELATIONSHIPS 1 to 1, 1 to Many
        self.valsvar7 = tk.StringVar()
        self.val7 = ("1:1", "1:N")
        self.cbox7 = ttk.Combobox(self.root, textvariable=self.valsvar7)
        self.cbox7['values'] = self.val7
        self.cbox7.current(0)
        self.cbox7['state'] = 'readonly'
        self.cbox7.bind('<<ComboboxSelected>>')
        self.valsvar7.set("1:1")
        self.cbox7.grid(column=1, row=r+3, padx=0, pady=5, sticky=tk.NW, columnspan=1)
        self.label7 = Label(self.root, text="Relationsips:")
        self.label7.grid(column=0, row=r+3, padx=0, pady=5, sticky=tk.NE, columnspan=1)

        self.var2SND1 = BooleanVar(value=True)
        self.switch2SND1 = ttk.Checkbutton(self.root, text="Output results to S&D Tool?", variable=self.var2SND1)
        self.switch2SND1.grid(column=1, row=r+4, sticky=tk.W, padx=10, pady=5)

        self.var2SND2 = BooleanVar(value=True)
        self.switch2SND2 = ttk.Checkbutton(self.root, text="Truncate Results?", variable=self.var2SND2)
        self.switch2SND2.grid(column=1, row=r+5, sticky=tk.W, padx=10, pady=5)

        self.SFBtn = tk.Button(self.root, text="Create SND RESPONSES/CONTRACTS File", command=lambda:
                                self.processData1(self.cbox1, self.cbox2, self.cbox7, self.fname1, self.fname2, topFldr, self.var2SND1, self.var2SND2), state=NORMAL, bg="green", fg="white")
        #self.SFBtn.grid(column=2, row=r+3, padx=10, pady=10, sticky=tk.W)
        self.SFBtn.grid(column=1, row=r+5, padx=200, pady=5, sticky=tk.W, columnspan=2)

        r=r+6

        self.fname3 = Text(self.root, width=66, height=1)
        self.fname3.insert("1.0", const.strNoFileDisp)
        self.fname3['state'] = 'disabled'
        self.fname3.grid(column=1, row=r, padx=10, pady=5, sticky=tk.W)
        self.fileBtn3 = ttk.Button(self.root, text="Select PSA SF file", command=lambda:
                                    self.fileBtn_clicked(self.fname3, '*PSA_SF-*', self.xlBtn3, False), state=NORMAL)
        self.fileBtn3.grid(column=0, row=r, padx=10, pady=5, sticky=tk.W)
        self.xlBtn3 = ttk.Button(self.root, text="Edit Data",
                                 command=lambda: ut.open_xl(self.fname3, self.dict_config[const.strWorkingFolder]),
                                 image=self.pic_excel_logo, compound=LEFT, state=DISABLED)
        self.xlBtn3.grid(column=2, row=r, padx=10, pady=0)
        
        #RESPONSES/CONTRACTS
        self.valsvar5 = tk.StringVar()
        self.val5 = ("RESPONSES", "CONTRACTS")
        self.cbox5 = ttk.Combobox(self.root, textvariable=self.valsvar5)
        self.cbox5['values'] = self.val5
        self.cbox5.current(0)
        self.cbox5['state'] = 'readonly'
        self.cbox5.bind('<<ComboboxSelected>>')
        self.valsvar5.set("RESPONSES")
        self.cbox5.grid(column=1, row=r+1, padx=0, pady=5, sticky=tk.NW, columnspan=1)
        self.label5 = Label(self.root, text="Output type:")
        self.label5.grid(column=0, row=r+1, padx=0, pady=5, sticky=tk.NE, columnspan=1)
        # CALCULATE HISTORICAL TRUE/FALSE
        self.valsvar6 = tk.StringVar()
        self.val6 = ("TRUE", "FALSE")
        self.cbox6 = ttk.Combobox(self.root, textvariable=self.valsvar6)
        self.cbox6['values'] = self.val6
        self.cbox6.current(0)
        self.cbox6['state'] = 'readonly'
        self.cbox6.bind('<<ComboboxSelected>>')
        self.valsvar6.set("TRUE")
        self.cbox6.grid(column=1, row=r+2, padx=0, pady=5, sticky=tk.NW, columnspan=1)
        self.label6 = Label(self.root, text="Calc historical:")
        self.label6.grid(column=0, row=r+2, padx=0, pady=5, sticky=tk.NE, columnspan=1)

        self.SFBtn = tk.Button(self.root, text="Create SND CANDIDATES Output File", command=lambda:
                                self.processData2(self.cbox5, self.cbox6, self.fname3, topFldr), state=NORMAL, bg="green", fg="white")
        self.SFBtn.grid(column=1, row=r+2, padx=200, pady=5, sticky=tk.W, columnspan=2)

        r= r+2+1 
        r = r + 1

        '''
        # Select Asset File for choosing Threshold 
        # This has been moved to its own section and so doesn't really need to be here anymore.
        self.fname9 = Text(self.root, width=66, height=1)
        self.fname9.insert("1.0", const.strNoFileDisp)
        self.fname9['state'] = 'disabled'
        self.fname9.grid(column=1, row=r, padx=10, pady=5, sticky=tk.W)
        self.fileBtn9 = ttk.Button(self.root, text="Select Asset_Data file", command=lambda:
                                    self.fileBtn_clicked(self.fname9, '*Asset*', self.xlBtn9, False), state=NORMAL)
        self.fileBtn9.grid(column=0, row=r, padx=10, pady=5, sticky=tk.W)
        self.xlBtn9 = ttk.Button(self.root, text="Edit Data",
                                 command=lambda: ut.open_xl(self.fname9, self.dict_config[const.strWorkingFolder]),
                                 image=self.pic_excel_logo, compound=LEFT, state=DISABLED)
        self.xlBtn9.grid(column=2, row=r, padx=10, pady=0)
        '''

        r = r + 1

        self.exitBtn = tk.Button(self.root, text="Exit", command=self.exitBtn_clicked, bg="black", fg="white")
        self.exitBtn.grid(column=1, row=r, columnspan=2, padx=200, pady=5, sticky=W)

        # Main Loop #
        self.root.transient(self)
        self.root.protocol("WM_DELETE_WINDOW", self.exitBtn_clicked)

    def validPSArunID(self, strRunID):
        valid = False
        if (strRunID[:3] == "AUT") or (strRunID[:3] == "MAN"):
            if len(strRunID) == 20:
                if len(strRunID.split("-")) == 5:
                    valid=True
        return valid

    def run_analysis_digest(self):
        if askyesno('Run Analysis Digest', 'Are you sure you want to run Analysis_Digest.ipynb?', parent=self.root):
            directory = f'{self.dict_config[const.strWorkingFolder]}{const.strDataResults}{self.cboxRuns.get().split("|")[0].rstrip()}'
            directory = directory.replace('/','\\')
            PSARunID = directory.split('\\')[-1]
            scenario = str(self.cbox.get())
            Jupyter_file = 'Analysis_Digest.ipynb'
            input_notebook = f'{self.dict_config[const.strWorkingFolder]}/packages/psa/{Jupyter_file}'

            #print(f'Scenario: {scenario}\nPSARunID: {PSARunID}\nInput_notebook: {input_notebook}\nDirectory: {directory}')
            subp.call(['python','PSA_Read_Digest.py', input_notebook, PSARunID, directory, scenario])
        return

    def cal_VHL_btn(self, fldr, topFldr, display):
        status = const.PSAok
        msg = "Event_Manager: cal_VHL_btn"
        self.strFldr = topFldr + str(fldr.get("1.0", "end").strip())
        self.singleRunAnalysis = False
        #extract last xxx_yyyy-mm-dd-hh-mm chars for single entry test
        PSArunID = self.strFldr[-20:]
        if self.validPSArunID(PSArunID):
            self.dir = self.strFldr
            self.strFldr = self.strFldr[:-20]
            self.ID = PSArunID
            display=False
            self.ID = PSArunID
            self.singleRunAnalysis = True
            tgt = PSArunID[:3]
            x = pa.VHL_overview(single=self.singleRunAnalysis, save=True, stopping_index=1, directory=self.dir, decorate=False, target=tgt)
            if len(x) == 0:
                # msg = "Invalid file : " + tempfile
                msg = "Invalid file"
                showwarning(title="WARNING", message=msg, parent=self.root)
            else:
                tempfile = x[0]
                self.dirs = dict(x[1])
        else: 
            # Get unput from textbox
            txtInput = self.textbox.get()
            # Convert txtInput to int
            numOfRuns = int(txtInput)
            if self.man_var.get() == 1 and self.aut_var.get() == 0:
                x = pa.VHL_overview(single=False, save=True, stopping_index=numOfRuns, directory=self.strFldr, decorate=True, target="MAN")
            elif self.aut_var.get() == 1 and self.man_var.get() == 0:
                x = pa.VHL_overview(single=False, save=True, stopping_index=numOfRuns, directory=self.strFldr, decorate=True, target="AUT")
            elif self.aut_var.get() == 1 and self.man_var.get() == 1:
                x = pa.VHL_overview(single=False, save=True, stopping_index=numOfRuns, directory=self.strFldr, decorate=True, target="both") 
            # Initiate variable path = to tuple x
            tempfile = x[0]
            self.dirs = dict(x[1])
        
        if status == const.PSAok:
            if os.path.isfile(tempfile):
                self.update_list()
                self.update_scenario_list(True)
                self.updateAssets()
                if display:
                    os.system('"' + tempfile + '"')
                #self.update_scenario_list(True)
            else:
                msg = "Invalid file : " + tempfile
                showwarning(title="WARNING", message=msg, parent=self.root)

    def updateAssets(self, event=None):
        value = self.cboxRuns.get()
        self.ID = value
        if value != "Invalid PSArunID data":
            # Split the value using the "|" as the delimiter
            parts = value.split("|")
            # Assign the first and second parts to separate variables
            first_part = parts[0]
            second_part = parts[1]
            run = first_part.rstrip()
            self.ID = str(run)
            self.dir = os.path.join(self.strFldr, run)
            s = str(self.cbox3.get())
            if s != "No scenario data":
                #tempfile, self.df_overview = pa.overview(PSARunID=self.ID, scenario=s, directory=self.dir, save=True, single=self.singleRunAnalysis)
                #self.assets = list(self.df_overview['Asset Name'])
                self.asset_dict = pa.get_asset_type(self.ID, self.dir)
                self.assets = pa.HL_overview(self.ID, directory=self.dir)[s]
                self.updateAssetList()
        return self.assets


    def HL_Overview(self, display):
        status = const.PSAok
        # Dir from dropdown
        # Get the value from the combo box
        value = self.cboxRuns.get()
        asset_name = str(self.cbox1.get())
        scenario = str(self.cbox3.get())

        if len(value) == 0 or value == "Invalid PSArunID data":
            msg = "No valid data exists for this PSArunID"
            showwarning(title="WARNING", message=msg, parent=self.root)
        else:
            # Split the value using the "|" as the delimiter
            parts = value.split("|")

            # Assign the first and second parts to separate variables
            first_part = parts[0]
            second_part = parts[1]

            run = first_part.rstrip()

            self.ID = str(run)

            self.dir = os.path.join(self.strFldr, run)
            s = str(self.cbox.get())
            if s == "No scenario data":     
                s = "BASE"

            if self.ovr_var.get() == 1:
                # Overview function
                frFile = self.dir + "\\" + self.ID + const.strPSAFlexReqts + ".xlsx"

                if os.path.isfile(frFile):
                    tempfile, self.df_overview = pa.overview(PSARunID=self.ID, scenario=s, directory=self.strFldr, save=True, single=self.singleRunAnalysis)
                    self.asset_dict = pa.get_asset_type(self.ID, self.dir)
                    self.assets = pa.HL_overview(self.ID, directory=self.dir)[s]
                else:
                    tempfile = ""
                    self.assets=["No assets to display"]
                    self.updateAssetList()

            elif self.per_var.get() == 1:
                # Performance function
                try:
                    tempfile = pa.plot_multiple_pl(['FEEDERS','RESULTS','LF_CALC'], PSARunID=self.ID, directory=self.strFldr, scenarios=s, save=True)
                except:
                    status = const.PSAfileExistError
                    msg = "Error viewing performance data for this PSArunID " + self.ID
                
            elif self.pltOut_var.get() == 1:
                # Plot outages From input files

                try:
                    tempfile = pa.plot_gantt_outages(PSARunID=self.ID, directory=self.strFldr, start_time='PSARunID', save=True, out_type=s)
                except BaseException as err:
                    status = const.PSAfileExistError
                    msg = f"Error viewing {s} outage data for this PSArunID " + self.ID
                    print(err)

            elif self.pltOut_var2.get() == 1:
                # Plot all asset activity files
                try:
                    tempfile = pa.plot_gantt_outages(PSARunID=self.ID, directory=self.strFldr, start_time='PSARunID', save=True, out_type='ALL', scenario=s)
                except BaseException as err:
                    status = const.PSAfileExistError
                    msg = f'Error viewing ALL outage data for this PSArunID' + self.ID
                    print(err)

            elif self.pltOut_var1.get() == 1:
                # Plot outages From Tables

                try:
                    tempfile = pa.plot_gantt_outages_debug(PSARunID=self.ID, directory=self.strFldr, out_type=s, save=True)
                except FileNotFoundError as err:
                    status = const.PSAfileExistError
                    msg = f'{err}'
                except BaseException as err:
                    status = const.PSAfileExistError
                    msg = "Error viewing CONT data for this PSArunID " + self.ID
                    print(err)

            elif self.pltEvt_var.get() == 1:
                # Plot Events function

                try:
                    tempfile = pa.plot_gantt_outages(PSARunID=self.ID, directory=self.strFldr, start_time='PSARunID', save=True, out_type='EVENT')
                except BaseException as err:
                    status = const.PSAfileExistError
                    msg = 'Error viewing EVENT data for this PSArunID ' + self.ID
                    print(err)

            elif self.sfAnlys_var.get() == 1:
                # Sensitivity Factors function
                try:
                    tempfile = pa.sf_analysis(PSARunID=self.ID, directory=self.strFldr, scenario=s, decorate=True, save=True)
                except BaseException as err:
                    status = const.PSAfileExistError
                    msg = err
                    print(err)

            elif self.pltThr_var.get() == 1:
                # Plot Thresholds
                try:
                    tempfile = pa.plot_asset_data(PSARunID=self.ID, directory=self.strFldr, relevant=True, save=True)
                except BaseException as err:
                    status = const.PSAfileExistError
                    msg = "Error viewing asset data for this PSARunID " + self.ID
                    print(err)

            if status == const.PSAok:
                if os.path.isfile(tempfile):
                    if display:
                        os.system('"' + tempfile + '"')
                        #self.updateAssetList() # This is unnecessary since the post command on the combobox for selecting assets should automatically do this step earlier
                        #self.update_scenario_list(True)
            else:
                showwarning(title="WARNING", message=msg, parent=self.root)


    def lowLevel_Overview(self):
        # get psa run ID
        # get scenario
        status = const.PSAok

        # Get asset name from drop down
        asset_name = str(self.cbox1.get())
        # Get PSA Run ID
        runID = str(self.ID)
        # Get Version
        try:
            batch_version = str(self.cbox4.get()).split(' | ')[1]
            batch, version = batch_version.split('-')
        except IndexError:
            version = 'latest'
            batch = 'latest'

        #indexed_df = self.df_overview.set_index('Asset Name')
        #asset_type = indexed_df.at[asset_name, 'Type']
        asset_type = self.asset_dict[asset_name]

        if runID == "Invalid PSArunID data":
            msg = "No valid data exists for this PSArunID"
            showwarning(title="WARNING", message=msg, parent=self.root)
        elif asset_name == "No assets to display":
            msg = "No valid data exists for this PSArunID"
            showwarning(title="WARNING", message=msg, parent=self.root)
        else:
            # Get scenario
            Scenario = str(self.cbox3.get())
            
            if self.Gra_var.get() == 1:
                # READ_ASSET_FUNCTION if graph checkbox checked then, set plot = True
                try:
                    tempfile = pa.read_asset(asset=asset_name, asset_type=asset_type, PSARunID=runID, directory=self.dir, scenario=Scenario, plot=True, save=True)
                except BaseException as err:
                    # change to open window notifty user folder does not exist 
                    status = const.PSAfileExistError
                    msg = "Error viewing data for PSArunID " + runID + " and Scenario " + Scenario
            elif self.hm_var.get() == 1:
                # READ_ASSET_FUNCTION if HEAT MAP checkbox checked then, set plot = False
                try:
                    tempfile = pa.read_asset(asset=asset_name, asset_type=asset_type, PSARunID=runID, directory=self.dir, scenario=Scenario, plot=False, save=True)
                except BaseException as err:
                    status = const.PSAfileExistError
                    msg = "Error viewing data for PSArunID " + runID + " and Scenario " + Scenario
            elif self.gra1_var.get() == 1:
                # READ FLEX RQ function if graph checkbox is checked then, set plot = False
                try:
                    tempfile = pa.read_flex_rq(asset=asset_name, PSARunID=runID, directory=self.dir, scenario=Scenario, plot=True, save=True)
                except BaseException as err:
                    status = const.PSAfileExistError
                    msg = "Error viewing data for PSArunID " + runID + " and Scenario " + Scenario 
            elif self.ex_var.get() == 1:
                # READ FLEX RQ function if excel checkbox is checked then, set plot = False
                try:
                    tempfile = pa.read_flex_rq(asset=asset_name, PSARunID=runID, directory=self.dir, scenario=Scenario, plot=False, save=True)
                except BaseException as err:
                    status = const.PSAfileExistError
                    msg = "Error viewing data for PSArunID " + runID + " and Scenario " + Scenario
            elif self.pltconres_var1.get() == 1:
                # Plot CONRES by batch
                try:
                    tempfile = pa.plot_multiple_conres_by_batch(PSARunID=self.ID, directory=self.strFldr, scenario=Scenario, batch=batch, save=True)
                except BaseException as err:
                    status = const.PSAfileExistError
                    msg = err
                    print(err)
            elif self.pltconres_var2.get() == 1:
                # Plot CONRES for a specific version
                try:
                    tempfile = pa.plot_conres(PSARunID=self.ID, directory=self.strFldr, scenario=Scenario, version_number=batch_version, save=True)
                except BaseException as err:
                    status = const.PSAfileExistError
                    msg = "Error viewing constraint resolution files for this PSARunID " + self.ID
                    print(err)
            
            if status == const.PSAok:
                if os.path.isfile(tempfile):
                    os.system('"' + tempfile + '"')
                else:
                    status = const.PSAfileExistError
                    msg = "Def Low_level_overview: File doesn't exist: " + tempfile
            
            if status != const.PSAok:
                showwarning(title="WARNING", message=msg, parent=self.root)

    def AM_tick_box(self):
        if self.aut_var.get() == 0 and self.man_var.get() == 0:
            self.aut_var.set(1)

    def AF_tick_box(self, chkBox):    
        self.ovr_var.set(0)
        self.per_var.set(0)
        self.pltOut_var.set(0)
        self.pltOut_var1.set(0)
        self.pltOut_var2.set(0)
        self.sfAnlys_var.set(0)
        self.pltEvt_var.set(0)
        self.pltThr_var.set(0)
        chkBox.set(1)

    def AF2_tick_box(self, chkBox):    
        self.Gra_var.set(0)
        self.hm_var.set(0)
        self.gra1_var.set(0)
        self.ex_var.set(0)
        self.pltconres_var1.set(0)
        self.pltconres_var2.set(0)
        chkBox.set(1)

    def AF3_tick_box(self, chkBox):
        self.gra2_var.set(0)
        self.gra3_var.set(0)
        chkBox.set(1)

    def update_list(self):
        list = pa.convert_to_dropdown(self.dirs)
        if len(list) > 0:
            self.cboxRuns.delete(0, "end")
            self.cboxRuns['values'] = list
            self.valsvar0.set(list[0])
            self.cboxRuns.current(0)
        else:
            list2 = []
            list2.append("Invalid PSArunID data")
            self.cboxRuns.delete(0, "end")
            self.cboxRuns['values'] = list2
            self.valsvar0.set(list2[0])
            self.cboxRuns.current(0)

    def update_scenario_list(self, bAssList):
        psarun = self.cboxRuns.get()
        #print(f'This is the psarun; {psarun}')
        scs = []
        if (len(psarun) == 0) or (psarun == "Invalid PSArunID data"):
            scs.append("No scenario data")
        else:
            if " BX" not in psarun:
                scs.append("BASE")
            if " MX" not in psarun:
                scs.append("MAINT")
            if " CX" not in psarun:
                scs.append("CONT")
            if " MCX" not in psarun:
                scs.append("MAINT_CONT")
            if (len(scs) == 0):
                scs.append("No scenario data")
        self.cbox.delete(0, "end")
        self.cbox['values'] = scs
        self.valsvar0.set(scs[0])
        self.cbox.current(0)
        if bAssList:
            self.cbox3.delete(0, "end")
            self.cbox3['values'] = scs
            #self.cbox3vals.set(scs[0])
            self.cbox3.current(0)


    def updateAssetList(self):
        list = self.assets
        self.cbox1.delete(0, "end")
        self.cbox1['values'] = list
        self.cbox1vals.set(list[0])
        self.cbox1.current(0)

    def event_test(self, event):
        self.update_scenario_list(True)
        self.assets=["No assets to display"]
        self.updateAssets()
        self.updateAssetList()
        self.update_events()
        self.update_conres()

    def update_conres(self):
        self.cbox4.delete(0, "end")
        try: 
            fldr = f'{self.strFldr}{self.ID}{const.strTempCRfiles}'
            values = [f'Version | {file[1:]}' for file in os.listdir(fldr)]
            values = tuple(values)
            self.cbox4['values'] = values
            self.cbox4.current(0)
            self.chkconres_var1['state']='normal'
            self.chkconres_var2['state']='normal'
        except BaseException as err:
            values = ("No_ConRes_Data")
            self.cbox4['values'] = values
            self.cbox4.current(0)
            self.chkconres_var1['state']='normal'
            self.chkconres_var2['state']='normal'
        return values

        
    # Analysis and reports window
    def __analysis(self):
        # ROOT WINDOW
        ws = 1050
        hs = 550
        self.root = tk.Toplevel(master=self, width=ws, height=hs)
        self.root.title("Analysis and Reporting")
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw/2) - (ws/2)
        y = (sh/2) - (hs/2)
        wstr = str(ws) + "x" + str(hs) + "+" + str(int(x)) + "+" + str(int(y))
        self.root.geometry(wstr)
        self.root.resizable(0,0)

        self.assets=["No assets to display"]
        self.val = ["No scenario data"]

        self.dict_config = ut.PSA_SND_read_config(const.strConfigFile)
        topFldr = self.dict_config[const.strWorkingFolder]
        self.pic_excel_logo = PhotoImage(file=self.dict_config[const.strWorkingFolder] + const.strExcelLogo)
        self.pic_excel_logo = self.pic_excel_logo.subsample(120)
        
        # Create lable frame lf0: Low Level
        self.lf0 = ttk.LabelFrame(self.root, text="HIGH LEVEL OVERVIEW")
        self.lf0.grid(column=0, row=0, padx=10, pady=5, sticky=tk.W, columnspan=2)

        # Select PSA folder button
        self.fldrBtn2 = ttk.Button(self.lf0, text="Select Folder for Analysis", command=lambda: self.folderBtn_clicked(self.fname1, '', self.fldrBtn2, False))
        self.fldrBtn2['state'] = 'NORMAL'
        self.fldrBtn2.grid(column=0, row=1, padx=10, pady=5, sticky=tk.W)
        topFldr = self.dict_config[const.strWorkingFolder]

        # box with file name inside 
        resultsFolder = const.strDataResults
        self.fname1 = Text(self.lf0, width=66, height=1)
        self.fname1.insert("1.0", resultsFolder)
        self.fname1['state'] = 'disabled'
        self.fname1.grid(column=1, row=1, padx=10, pady=5, sticky=tk.W)

        # Automatic or Manual lable 
        self.VHLlbl = Label(self.lf0, text="Select Automatic or Manual runs")
        self.VHLlbl.grid(column=0, row=2, padx=10, pady=5)

        # Checkbutton for manual runs
        self.man_var = IntVar()
        self.man_var.set(1)
        self.chkMan = ttk.Checkbutton(self.lf0, text = "Manual", variable=self.man_var, command= lambda: self.AM_tick_box())
        self.chkMan.grid(column=0, row=3, padx=10, pady=5, sticky=tk.W)

        # Checkbutton for automatic runs
        self.aut_var = IntVar()
        self.aut_var.set(1)
        self.chkAut = ttk.Checkbutton(self.lf0, text = "Automatic", variable=self.aut_var, command= lambda: self.AM_tick_box())
        self.chkAut.grid(column=0, row=4, padx=10, pady=5, sticky=tk.W)

        # Number of runs
        self.numRunsLbl = Label(self.lf0, text="Number of runs to calculate:")
        self.numRunsLbl.grid(column=1, row=2, padx=5, pady=5, sticky=tk.W)

        # create a textbox
        self.textbox = Entry(self.lf0, text="ENTER")
        # Clear textbox before enter number
        self.textbox.delete(0, "end")
        self.textbox.insert(0, int("6"))
        self.textbox.grid(column=1, row=3, padx=10, pady=5, sticky=tk.W)

        # Display results in excel button 
        self.disBtn1 = ttk.Button(self.lf0, text="Update PSArunIDs",
                                 command=lambda: self.cal_VHL_btn(self.fname1, topFldr, False),
                                 image=self.pic_excel_logo, compound=LEFT, state=NORMAL)
        self.disBtn1.grid(column=2, row=1, padx=20, pady=5)
        
        # Create lable frame lf1
        self.lf1 = ttk.LabelFrame(self.root, text="PSArunID LEVEL OVERVIEW")
        self.lf1.grid(column=0, row=1, padx=10, pady=5, sticky=tk.W, columnspan=2)

        # Select Runs label
        self.sltLbl = Label(self.lf1, text="Select PSArunID from Dropdown Menu")
        self.sltLbl.grid(column=0, row=0, padx=10, pady=5, stick=tk.W)

        # Drop down box for selecting runs
        self.cboxRuns = ttk.Combobox(self.lf1, width = 38)
        self.valsvar0 = tk.StringVar()
        self.cboxRuns['state'] = 'readonly'
        self.cboxRuns.bind('<<ComboboxSelected>>', self.event_test)
        self.cboxRuns.grid(column=0, row=1, padx=10, pady=5, sticky=tk.W, columnspan=2)
        
        # Select Scenario
        self.sltLbl = Label(self.lf1, text="Select Scenario from Dropdown")
        self.sltLbl.grid(column=0, row=2, padx=10, pady=5, sticky=tk.W)

        # Dropdown for Scenarios
        self.cbox = ttk.Combobox(self.lf1, width = 15)
        self.valsvar0 = tk.StringVar()
        self.cbox['values'] = self.val
        self.cbox.current(0)
        self.cbox['state'] = 'readonly'
        self.cbox.bind('<<ComboboxSelected>>')
        self.cbox.grid(column=0, row=3, padx=10, pady=5, sticky=tk.W, columnspan=2)

        # Analysis Functions Label 
        self.aFlbl = Label(self.lf1, text="Analysis Functions")
        self.aFlbl.grid(column=1, row=0, padx=10, pady=5)

        # Checkbutton for Overview
        self.ovr_var = IntVar()
        self.ovr_var.set(1)
        self.chkOvr = ttk.Checkbutton(self.lf1, text = "Overview", variable=self.ovr_var,
                                    command= lambda: self.AF_tick_box(self.ovr_var))
        self.chkOvr.grid(column=1, row=1, padx=55, pady=5, sticky=tk.W)

        # Checkbutton for Performance
        self.per_var = IntVar()
        self.per_var.set(0)
        self.chkPer = ttk.Checkbutton(self.lf1, text = "Performance", variable=self.per_var,
                                    command= lambda: self.AF_tick_box(self.per_var))
        self.chkPer.grid(column=1, row=2, padx=55, pady=5, sticky=tk.W)

        # Checkbutton for Sensitivity Factors Analysis
        self.sfAnlys_var = IntVar()
        self.sfAnlys_var.set(0)
        self.chkSF = ttk.Checkbutton(self.lf1, text = "Sensitivity Factors", variable=self.sfAnlys_var, 
                                    command= lambda: self.AF_tick_box(self.sfAnlys_var))
        self.chkSF.grid(column=1, row=3, padx=55, pady=5, sticky=tk.W)

        # Checkbutton for Plot Asset Thresholds
        self.pltThr_var = IntVar()
        self.pltThr_var.set(0)
        self.chkThr = ttk.Checkbutton(self.lf1, text = "Plot Asset Thresholds", variable=self.pltThr_var,
                                    command= lambda: self.AF_tick_box(self.pltThr_var))
        self.chkThr.grid(column=1, row=4, padx=55, pady=5, sticky=tk.W)

        # Checkbutton for Plot Inputs
        self.pltOut_var = IntVar()
        self.pltOut_var.set(0)
        self.chkOut = ttk.Checkbutton(self.lf1, text = "Plot Outages from Inputs", variable=self.pltOut_var,
                                    command= lambda: self.AF_tick_box(self.pltOut_var))
        self.chkOut.grid(column=2, row=1, padx=30, pady=5, sticky=tk.W)

        # Checkbutton for Plot Tables
        self.pltOut_var1 = IntVar()
        self.pltOut_var1.set(0)
        self.chkOut1 = ttk.Checkbutton(self.lf1, text = "Plot Outages from Tables", variable=self.pltOut_var1,
                                    command= lambda: self.AF_tick_box(self.pltOut_var1))
        self.chkOut1.grid(column=2, row=2, padx=30, pady=5, sticky=tk.W)

        # Checkbutton for Plot Events
        self.pltEvt_var = IntVar()
        self.pltEvt_var.set(0)
        self.chkEvt = ttk.Checkbutton(self.lf1, text = "Plot Events", variable=self.pltEvt_var,
                                    command= lambda: self.AF_tick_box(self.pltEvt_var))
        self.chkEvt.grid(column=2, row=3, padx=30, pady=5, sticky=tk.W)

        # Checkbutton for Plot All activity
        self.pltOut_var2 = IntVar()
        self.pltOut_var2.set(0)
        self.chkOut2 = ttk.Checkbutton(self.lf1, text = "Plot All Asset Activity", variable=self.pltOut_var2,
                                    command= lambda: self.AF_tick_box(self.pltOut_var2))
        self.chkOut2.grid(column=2, row=4, padx=30, pady=5, sticky=tk.W)
        
        # Run Analysis Digest button
        self.digestBtn1 = ttk.Button(self.lf1, text='Run Analysis Digest', 
                                 command=self.run_analysis_digest)
        self.digestBtn1.grid(column=3, row=4, padx=5, pady=0, sticky=tk.E)

        # Display HL results button 
        self.disBtn1 = ttk.Button(self.lf1, text="Display results and update AssetIDs",
                                 command=lambda: self.HL_Overview(True), state=NORMAL)
        self.disBtn1.grid(column=3, row=0, padx=5, pady=5)
 
        # Create lable frame lf2: Low Level
        self.lf2 = ttk.LabelFrame(self.root, text="ASSET LEVEL OVERVIEW")
        self.lf2.grid(column=0, row=2, padx=10, pady=5, sticky=tk.W)

        # Select scenario label
        self.assLbl = Label(self.lf2, text="Select AssetID to Analyse")
        self.assLbl.grid(column=0, row=0, padx=10, pady=5, sticky=tk.W)

        # Drop down box for selecting runs
        self.cbox1 = ttk.Combobox(self.lf2, width = 25) #postcommand=lambda: self.cbox1.configure(values=self.updateAssets())
        self.cbox1vals = tk.StringVar()
        self.cbox1['state'] = 'readonly'
        self.cbox1.bind('<<ComboboxSelected>>')
        self.cbox1.grid(column=0, row=1, padx=10, pady=5, sticky=tk.W)

        # Select scenario label
        self.sceLbl = Label(self.lf2, text="Select Scenario from Dropdown")
        self.sceLbl.grid(column=0, row=2, padx=10, pady=5, sticky=tk.W)

        # Dropdown for Scenarios
        self.cbox3 = ttk.Combobox(self.lf2, width = 15)
        #self.cbox3vals = tk.StringVar()
        self.cbox3['values'] = self.val
        self.cbox3.current(0)
        self.cbox3['state'] = 'readonly'
        self.cbox3.bind('<<ComboboxSelected>>', self.updateAssets)
        self.cbox3.grid(column=0, row=3, padx=10, pady=5, sticky=tk.W, columnspan=2)

        if self.autorun_var.get():
            self.cal_VHL_btn(self.fname1, topFldr, False)
            self.cboxRuns.current(0)
            #self.update_scenario_list(False)
            #self.HL_Overview(False)
            self.cbox1.current(0)
        else:
            self.updateAssetList()        
        self.update_scenario_list(True)

        # Analysis Functions Label 
        self.anFlbl = Label(self.lf2, text="Analysis Functions")
        self.anFlbl.grid(column=1, row=0, padx=10, pady=5, sticky=tk.EW, columnspan=2)

        # % Loading label
        self.rALbl = Label(self.lf2, text="% " +"Loading")
        self.rALbl.grid(column=1, row=1, padx=10, pady=5, sticky=tk.W)

        # Checkbutton for Graph
        self.Gra_var = IntVar()
        self.Gra_var.set(0)
        self.chkGra = ttk.Checkbutton(self.lf2, text = "Graph", variable=self.Gra_var,
                                            command= lambda: self.AF2_tick_box(self.Gra_var))
        self.chkGra.grid(column=1, row=2, padx=10, pady=5, sticky=tk.W)

        # Checkbutton for Heat Map
        self.hm_var = IntVar()
        self.hm_var.set(1)
        self.chkHM = ttk.Checkbutton(self.lf2, text = "Heat Map", variable=self.hm_var,
                                                    command= lambda: self.AF2_tick_box(self.hm_var))
        self.chkHM.grid(column=1, row=3, padx=10, pady=5, sticky=tk.W)

        # Required kW Label
        self.rfrqLbl = Label(self.lf2, text="Required kW")
        self.rfrqLbl.grid(column=2, row=1, padx=10, pady=5, sticky=tk.W)

        # Checkbutton for Graph
        self.gra1_var = IntVar()
        self.gra1_var.set(0)
        self.chkGra1 = ttk.Checkbutton(self.lf2, text = "Graph", variable=self.gra1_var,
                                                    command= lambda: self.AF2_tick_box(self.gra1_var))
        self.chkGra1.grid(column=2, row=2, padx=10, pady=5, sticky=tk.W)

        # Checkbutton for Excel File
        self.ex_var = IntVar()
        self.ex_var.set(0)
        self.chkEx = ttk.Checkbutton(self.lf2, text = "Excel File", variable=self.ex_var,
                                                    command= lambda: self.AF2_tick_box(self.ex_var))
        self.chkEx.grid(column=2, row=3, padx=10, pady=5, sticky=tk.W)

        # Constraint Resolution Label
        self.rfrqLbl = Label(self.lf2, text="Constraint Resolution")
        self.rfrqLbl.grid(column=3, row=0, padx=10, pady=5, sticky=tk.W)

        # Checbutton for Conres by version (Batch Level)
        self.pltconres_var1 = IntVar()
        self.pltconres_var1.set(0)
        self.chkconres_var1 = ttk.Checkbutton(self.lf2, text = "Batch Level", variable=self.pltconres_var1,
                                                    command= lambda: self.AF2_tick_box(self.pltconres_var1))
        self.chkconres_var1.grid(column=3, row=2, padx=10, pady=5, sticky=tk.W)

        # Checkbutton for Conres explicit (Version Level)
        self.pltconres_var2 = IntVar()
        self.pltconres_var2.set(0)
        self.chkconres_var2 = ttk.Checkbutton(self.lf2, text = "Version Level", variable=self.pltconres_var2,
                                                    command= lambda: self.AF2_tick_box(self.pltconres_var2))
        self.chkconres_var2.grid(column=3, row=3, padx=10, pady=5, sticky=tk.W)

        # Select Version Combobox
        self.cbox4 = ttk.Combobox(self.lf2, width = 15)
        self.cbox4vals = tk.StringVar()
        #self.cbox4val = ("No_ConRes_Data")
        #self.valsvar4 = self.cbox4val
        self.cbox4['values'] = self.update_conres()
        self.cbox4.current(0)
        self.cbox4['state'] = 'readonly'
        self.cbox4.bind('<<ComboboxSelected>>')
        self.cbox4.grid(column=3, row=1, padx=10, pady=5, sticky=tk.W)

        # Display LL
        self.disLLBtn1 = ttk.Button(self.lf2, text="Display results",
                                 command=lambda: self.lowLevel_Overview(), state=NORMAL)
        self.disLLBtn1.grid(column=4, row=0, padx=5, pady=5, sticky=tk.W)

        self.exitBtn = tk.Button(self.root, text="Exit", command=self.exitBtn_clicked, bg="black", fg="white")
        self.exitBtn.grid(column=0, row=5, padx=5, pady=5, columnspan=2)

        # Create lable frame lf3: Low Level
        self.lf3 = ttk.LabelFrame(self.root, text="GENERATOR LEVEL OVERVIEW")
        self.lf3.grid(column=1, row=2, padx=10, pady=5, sticky=tk.NSEW)

        # Select Event Asset
        self.event_label = tk.Label(self.lf3, text='Select Event Asset')
        self.event_label.grid(column=0, row=0, padx=5, pady=5, sticky=tk.W)

        self.event_list = ('Working')
        self.event_var = tk.StringVar()
        self.cbox10 = ttk.Combobox(self.lf3, values=self.event_list, textvariable=self.event_var, width = 30)
        self.cbox10['state'] = 'readonly'
        self.cbox10.grid(column=0, row=1, padx=10, pady=5, sticky=tk.W)

        self.function_label = tk.Label(self.lf3, text='Power Dispatch')
        self.function_label.grid(column=0, row=2, padx=10, stick=tk.W)

        # Checkbutton for Graph
        self.gra2_var = IntVar()
        self.gra2_var.set(1)
        self.chkGra2 = ttk.Checkbutton(self.lf3, text = "Graph", variable=self.gra2_var,
                                                    command= lambda: self.AF3_tick_box(self.gra2_var))
        self.chkGra2.grid(column=0, row=3, padx=10, pady=5, sticky=tk.W)

        # Checkbutton for Graph
        self.gra3_var = IntVar()
        self.gra3_var.set(0)
        self.chkGra3 = ttk.Checkbutton(self.lf3, text = "Excel File", variable=self.gra3_var,
                                                    command= lambda: self.AF3_tick_box(self.gra3_var))
        self.chkGra3.grid(column=0, row=3, padx=10, pady=5, sticky=tk.E)

        # Display LL
        self.disLLBtn2 = ttk.Button(self.lf3, text="Display results",
                                 command=lambda: self.eventLevel_overview(), state=NORMAL)
        self.disLLBtn2.grid(column=2, row=0, padx=5, pady=5, sticky=tk.W)

        self.cbox10['values'] = self.update_events()
        self.cbox10.current(0)
        
        # Main Loop #
        self.root.transient(self)
        self.root.protocol("WM_DELETE_WINDOW", self.exitBtn_clicked)

    def eventLevel_overview(self):
        # get psa run ID
        # get scenario
        status = const.PSAok

        # Get asset name from drop down
        asset_name = str(self.cbox10.get())
        # Get PSA Run ID
        runID = str(self.ID)

        if runID == "Invalid PSArunID data":
            msg = "No valid data exists for this PSArunID"
            showwarning(title="WARNING", message=msg, parent=self.root)
        elif asset_name == "No Events Found":
            msg = "No valid data exists for this PSArunID"
            showwarning(title="WARNING", message=msg, parent=self.root)
        else:
            # Get scenario
            Scenario = str(self.cbox3.get())
            
            if self.gra2_var.get() == 1:
                # Show plot
                try:
                    tempfile = pa.plot_p_dispatch(asset=asset_name, PSARunID=runID, directory=self.dir, save=True)
                except BaseException as err:
                    # change to open window notifty user folder does not exist
                    print(err) 
                    status = const.PSAfileExistError
                    msg = "Error viewing data for PSArunID " + runID + " and Scenario " + Scenario 
            elif self.gra3_var.get() == 1:
                # Show Data
                try:
                    tempfile = f'{self.dir}{const.strEventFiles}{asset_name}.xlsx'
                except:
                    status = const.PSAfileExistError
                    msg = "Error viewing data for PSArunID " + runID + " and Scenario " + Scenario 
                
            if status == const.PSAok:
                if os.path.isfile(tempfile):
                    os.system('"' + tempfile + '"')
                else:
                    status = const.PSAfileExistError
                    msg = "Def Low_level_overview: File doesn't exist: " + tempfile
            
            if status != const.PSAok:
                showwarning(title="WARNING", message=msg, parent=self.root)

    def update_events(self):
        self.cbox10.delete(0, "end")
        try: 
            fldr = f'{self.strFldr}{self.ID}{const.strEventFiles}'
            values = [file[:-5] for file in os.listdir(fldr)]
            values = tuple(values)
            self.cbox10['values'] = values
            self.cbox10.current(0)
            self.disLLBtn2['state']='normal'
        except:
            values = ("No_Event_Data")
            self.cbox10['values'] = values
            self.cbox10.current(0)
            self.disLLBtn2['state']='disabled'
        return values


if __name__ == '__main__':
    App().mainloop()
