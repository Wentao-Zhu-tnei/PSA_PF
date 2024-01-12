import os
import pandas as pd
import csv
import time
import PF_Config as pfconf
import PSA_SND_Utilities as ut
import PSA_SND_Constants as const
import PSA_PF_Functions as pffun
import PSA_File_Validation as fv
import PSA_SIA_Data_Input_Output as sio
import PSA_NeRDA_Data_Input_Output as nrd
import PSA_ProgressBar as pbar
import powerfactory_interface as pf_interface
from datetime import datetime

def PSA_filter_NeRDA(dfNeRDACopy,dfAssetNeRDA):
    try:
        lNeRDA_asset_required = dfAssetNeRDA['PF_ID'].to_list()
    except:
        lNeRDA_asset_required = dfAssetNeRDA['ASSET_LONG_NAME'].to_list()
    for index, row in dfNeRDACopy.iterrows():
        if row['shortName'] not in lNeRDA_asset_required:
            dfNeRDACopy = dfNeRDACopy.drop(index)
    return dfNeRDACopy

def PSA_update_switch(pf_object, dfElmCoup_ALL, dfStaSwitch_ALL, dfNeRDA):
    """
    Find the switch/breaker from NeRDA output and update the status of corresponding component in PF
    """

    for index, row in dfNeRDA.iterrows():
        switch = row['shortName']
        status = row['status']
        if status == "OPEN":
            switch_status = 0
            switch_action = 'open'
        elif status == "CLOSED":
            switch_status = 1
            switch_action = 'close'
        else:
            switch_status = -1
            print(f"The status of switch [{switch}] from NeRDA is unknown")
        if switch_status != -1:
            if switch in dfElmCoup_ALL['Name'].to_list():
                strObjSwitch = dfElmCoup_ALL[dfElmCoup_ALL['Name'] == switch]['PF_Object_Coupler'].iloc[0]
                if strObjSwitch.isclosed != switch_status:  # returns 0 or 1
                    pf_object.toggle_switch_coup_ind(strObjSwitch,action=switch_action,verbose=pfconf.bDebug)

            elif switch in dfStaSwitch_ALL['Name'].to_list():
                strObjSwitch = dfStaSwitch_ALL[dfStaSwitch_ALL['Name'] == switch]['PF_Object_Switches'].iloc[0]
                if strObjSwitch.isclosed != switch_status:  # returns 0 or 1
                    pf_object.toggle_switch_ind(strObjSwitch, action=switch_action, verbose=pfconf.bDebug)

            else:
                if pfconf.bDebug:
                    print(f"[{switch}] is not part of the PowerFactory Cowley Local BSP model")
    return

def PSA_calc_flex_reqts(isAuto, bOutput, isBase, PSAFolder, PSArunID, bEvents, eventsFile, plout, ploutfile, unplout, unploutfile, 
                        bSwitch, switchFile, lne_thresh,
                        trafo_thresh, power_factor, lagging, start_time_row, bDebug, PSA_running_mode, pf_proj_dir,
                        start_time, dictSIAFeeder, dictSIAGen, dictSIAGroup, NeRDA_output_path, lNamesGrid, dfDownScalingFactor_grid,
                        dfAllTrafos_grid_PQnom, NeRDAfolder, SIA_folder, strScenario_folder, input_folder,
                        runNum, numRuns):
    """ """
    # Function takes two arguments: PSAFolder, PSArunID
    # TODO: list of functions to be added:
    # Read parameters file into a dict
    # Read asset data file if it exists into a df
    # Read planned outages data file if it exists into a df
    # Read unplanned outages data file if it exists into a df
    # Load NERDA data
    # Output NERDA data to file (switch will be OPEN or CLOSED or NULL)
    # Load SIA data
    # Output SIA data to file
    # Return status and msg

    print("\n\n################################################################")
    print(f"SIMULATION STARTED at [{datetime.now()}]")
    print(f'Running PSA scenario: [{PSA_running_mode}]\n')
    print("################################################################\n\n")

    strTitle = "Run " + str(runNum) + " of " + str(numRuns) + " : " + PSArunID
    bBtn = (not isAuto)
    pbar.createProgressBar(strTitle, bBtn)
    pb=0
    pbar.updateProgressBar(pb, "Starting Process")

    status = const.PSAok
    msg = 'Calc_Flex_Reqts'

# Record performance time stamp
    header = ["CONFIG", "ASSETS", "PF_START", "NeRDA", "SIA", "OUTAGES", "PREP_LF", 
            "ALL_ASSETS", "CALC_LOADING", "OVERLOAD1", "OVERLOAD2", "COMBINED", "OUTPUT_FILES", "PF_EXIT"]

    if PSA_running_mode == "B":
        scenario = "BASE"
        output_path = strScenario_folder + const.strBfolder
    elif PSA_running_mode == "M":
        scenario = "MAINT"
        output_path = strScenario_folder + const.strMfolder
    elif PSA_running_mode == "C":
        scenario = "CONT"
        output_path = strScenario_folder + const.strCfolder
    elif PSA_running_mode == "MC":
        scenario = "MAINT_CONT"
        output_path = strScenario_folder + const.strMCfolder
    else:
        scenario = "UNKNOWN"
        status = const.PSAfileTypeError
        msg = "Calc_Flex_Reqts: Invalid PSA running mode = " + PSA_running_mode
    
    perf_file = os.path.dirname(os.path.dirname(output_path)) + '\\' + PSA_running_mode + '_performance_log2.csv'
    with open(perf_file, 'w', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(header)
        f.close()
    t0 = time.time()

    end_time = datetime.now()
    dfCombinedEvents = pd.DataFrame()
    dfCombinedConstraints = pd.DataFrame()
    filepath = PSAFolder + "\\"
    dict_config = ut.PSA_SND_read_config(const.strConfigFile)
    dict_params = ut.readPSAparams(filepath + PSArunID + const.strRunTimeParams)
    print(f'[const.strPFUser]={dict_config[const.strPFUser]}')
    print(f'[const.strBSPModel]={dict_params[const.strBSPModel]}')
    t1 = time.time()
    
    pb=2
    pbar.updateProgressBar(pb, "Preparing Data")

    #Extract the date element YYYY-MM-DD from PSArunID XXX_YYYY-MM-DD-hh-mm
    start_date = PSArunID[4:14]
    start_time = datetime.strptime(start_date + " 00:00:00", '%Y-%m-%d %H:%M:%S')
    fname_BSPmodel = dict_params[const.strBSPModel]
    fname_AssetData = dict_params[const.strAssetData]
    fname_EventData = dict_params[const.strEventFile]
    fname_ploutData = dict_params[const.strMaintFile]
    fname_unploutData = dict_params[const.strContFile]
    fname_SwitchData = dict_params[const.strSwitchFile]
    number_of_days = int(dict_params[const.strDays])
    number_of_half_hours = int(dict_params[const.strHalfHours])
    dfAssetData = pd.DataFrame()
    dfSIAData = pd.DataFrame()
    dfNeRDAData = pd.DataFrame()
    # dfPloutData = pd.DataFrame()
    # dfUnploutData = pd.DataFrame()
    # Read asset, plout, unplout data
    for r, d, f in os.walk(input_folder):
        for file in f:
            if file == fname_AssetData:
                dfAssetData = pd.read_excel(input_folder + "\\" + file, sheet_name=const.strAssetSheet)
                dfSIAData = pd.read_excel(input_folder + "\\" + file, sheet_name=const.strSIASheet)
                dfNeRDAData = pd.read_excel(input_folder + "\\" + file, sheet_name=const.strNeRDASheet)
                # print(dfAssetData)
            # elif file == fname_ploutData:
            #     dfPloutData = pd.read_excel(filepath + file, sheet_name=const.strSheetName)
            #     # print(dfPloutData)
            # elif file == fname_unploutData:
            #     dfUnploutData = pd.read_excel(filepath + file, sheet_name=const.strSheetName)
            #     # print(dfUnploutData)
    t2 = time.time()

    # Load NeRDA data
    print(f'[const.strPFUser]={dict_config[const.strPFUser]}')
    pf_object, status, msg = pf_interface.run(dict_config[const.strPFUser], False)

    if status != const.PSAok:
        pbar.updateProgressBar(100, "Error reading input data files")
        return status, msg, end_time, dictSIAFeeder, dictSIAGen, dictSIAGroup, start_time, NeRDA_output_path, \
           lNamesGrid, dfDownScalingFactor_grid,dfAllTrafos_grid_PQnom, dfCombinedEvents

    pb=3
    pbar.updateProgressBar(pb, "Importing Power Factory Model")

    pf_interface.activate_project(pf_object, fname_BSPmodel, pf_proj_dir)
    t3 = time.time()

    pb=4
    pbar.updateProgressBar(pb, "Power Factory Started")

    if isBase and pfconf.bUseNeRDA:
        # NeRDA_output_path = nrd.createNeRDAfolder(NeRDAfoldernumbering,PSAFolder)
        # NeRDA_output_path = filepath + "NeRDA_DATA/" + PSArunID
    
        pb=10
        pbar.updateProgressBar(pb, "Loading NeRDA Data")

        status, msg, dfNeRDA = nrd.loadNeRDAData(assets=dfNeRDAData, output_path=NeRDAfolder + "\\" + PSArunID,
                                        user=dict_config[const.strNERDAUsr],pwd=dict_config[const.strNERDAPwd])
        if status == const.PSAok and pfconf.bUseNeRDA:
        
            pb=pb+1
            pbar.updateProgressBar(pb, "Processing NeRDA Data")
        
            print("=" * 65 + "\n" + "Loading NeRDA data" + "\n" + "=" * 65)
            dfNeRDACopy = dfNeRDA.copy()
            dfNeRDACopy = PSA_filter_NeRDA(dfNeRDACopy, dfNeRDAData)

            pb=pb+1
            pbar.updateProgressBar(pb, "Processing Circuit Breakers")

            dfElmCoup_ALL = pf_object.get_all_circuit_breakers_coup()

            pb=pb+1
            pbar.updateProgressBar(pb, "Processing Switches")

            dfStaSwitch_ALL = pf_object.get_all_circuit_breakers_switch()

            pb=pb+1
            pbar.updateProgressBar(pb, "Updating Switch Positions")

            PSA_update_switch(pf_object, dfElmCoup_ALL, dfStaSwitch_ALL, dfNeRDACopy)

            pb=pb+1
            pbar.updateProgressBar(pb, "Recording Before/After Switch Positions")

            dfElmCoup_ALL_after = pf_object.get_all_circuit_breakers_coup()
            dfStaSwitch_ALL_after = pf_object.get_all_circuit_breakers_switch()

            pb=pb+1
            pbar.updateProgressBar(pb, "Saving NeRDA Data")

            ut.out_xl(dfs=[dfElmCoup_ALL.drop(columns='PF_Object_Coupler'),
                           dfStaSwitch_ALL.drop(columns='PF_Object_Switches'),
                           dfElmCoup_ALL_after.drop(columns='PF_Object_Coupler'),
                           dfStaSwitch_ALL_after.drop(columns='PF_Object_Switches')],
                      sheet_names=['Coupler_before', 'Switch_before', 'Coupler_after', 'Switch_after'],
                      outpath=os.path.join(NeRDAfolder, 'switch_status.xlsx'), engine='xlsxwriter')

    if status != const.PSAok:
        pbar.updateProgressBar(100, "Error saving NeRDA data")
        return status, msg, end_time, dictSIAFeeder, dictSIAGen, dictSIAGroup, start_time, NeRDA_output_path, \
           lNamesGrid, dfDownScalingFactor_grid,dfAllTrafos_grid_PQnom, dfCombinedEvents
    else:
        if not isBase and pfconf.bUseNeRDA:
            print("=" * 65 + "\n" + "Loading NeRDA data" + "\n" + "=" * 65)
            print("################")
            dfNeRDA = pd.read_excel(NeRDAfolder + "\\" + PSArunID +'-NeRDA_DATA.xlsx')
            # dfNeRDA = pd.read_excel(os.path.join(NeRDAfolder, PSArunID + '-NeRDA_DATA.xlsx'))
            dfNeRDACopy = dfNeRDA.copy()
            dfNeRDACopy = PSA_filter_NeRDA(dfNeRDACopy, dfNeRDAData)
            dfElmCoup_ALL = pf_object.get_all_circuit_breakers_coup()
            dfStaSwitch_ALL = pf_object.get_all_circuit_breakers_switch()
            PSA_update_switch(pf_object, dfElmCoup_ALL, dfStaSwitch_ALL, dfNeRDACopy)
            dfElmCoup_ALL_after = pf_object.get_all_circuit_breakers_coup()
            dfStaSwitch_ALL_after = pf_object.get_all_circuit_breakers_switch()
            ut.out_xl(dfs=[dfElmCoup_ALL.drop(columns='PF_Object_Coupler'),
                           dfStaSwitch_ALL.drop(columns='PF_Object_Switches'),
                           dfElmCoup_ALL_after.drop(columns='PF_Object_Coupler'),
                           dfStaSwitch_ALL_after.drop(columns='PF_Object_Switches')],
                      sheet_names=['Coupler_before', 'Switch_before', 'Coupler_after', 'Switch_after'],
                      outpath=os.path.join(NeRDAfolder, 'switch_status.xlsx'), engine='xlsxwriter')
        t4 = time.time()
        # Load SIA data
        # output_path = filepath + PSArunID

        pb=pb+1
        pbar.updateProgressBar(pb, "Loading SIA Forecast Data")

        token, status, msg = sio.SIA_get_token(dict_config[const.strSIAUsr], dict_config[const.strSIAPwd],
                                               int(dict_config[const.strSIATimeOut]))
        if status != const.PSAok:
            pbar.updateProgressBar(100, "Error loading SIA data")
            return status, msg, end_time, dictSIAFeeder, dictSIAGen, dictSIAGroup, start_time, NeRDA_output_path, \
           lNamesGrid, dfDownScalingFactor_grid,dfAllTrafos_grid_PQnom, dfCombinedEvents
        else:
            # SIA_folder = sio.createSIAfolder(PSAFolder)
            if isBase:
                start_time_SIA, start_datetime, status, msg, dictSIAFeeder, dictSIAGen, dictSIAGroup = sio.loadSIAData(
                    strbUTC=dict_config[const.bUTC], assets=dfSIAData, timeout=int(dict_config[const.strSIATimeOut]),
                    OutputFilePath=SIA_folder, token=token, bDebug=bDebug,PSArunID=PSArunID,
                    number_of_days=number_of_days, number_of_half_hours=number_of_half_hours)
            t5 = time.time()

            pb=pb+1
            pbar.updateProgressBar(pb, "Updating MAINT and CONT Outage Data")

            if status != const.PSAok:
                pbar.updateProgressBar(100, "Error loading SIA data")
                return status, msg, end_time, dictSIAFeeder, dictSIAGen, dictSIAGroup, start_time, NeRDA_output_path, \
                       lNamesGrid, dfDownScalingFactor_grid,dfAllTrafos_grid_PQnom, dfCombinedEvents
            else:
                if unplout:
                    dfUnplout = pffun.PF_update_cont(unploutfile,PSArunID, number_of_days)
                else:
                    dfUnplout = pd.DataFrame()
                if plout or unplout:
                    df_outages = pffun.PF_combineOutages(bDebug, plout, unplout, ploutfile, "DATA", dfUnplout, input_folder)
                    lOutageAssets, lOutageAssetTypes, lDfOutages = pffun.PF_mapOutages(dict_config[const.bUTC], start_time, df_outages,
                                                                                       input_folder + "\\" + PSA_running_mode + "-ASSET_OUTAGE_TABLES",
                                                                                       number_of_days, number_of_half_hours)
                else:
                    lOutageAssets, lOutageAssetTypes, lDfOutages = (pd.DataFrame() for i in range(3))
                t6 = time.time()
                pf_object.prepare_loadflow(ldf_mode=dict_config[const.strPFLFMode], algorithm=dict_config[const.strPFAlgrthm],
                     trafo_tap=dict_config[const.strPFTrafoTap], shunt_tap=dict_config[const.strPFShuntTap],
                     zip_load=dict_config[const.strPFZipLoad], q_limits=dict_config[const.strPFQLim],
                     phase_shift=dict_config[const.strPFPhaseShift],trafo_tap_limit=dict_config[const.strPFTrafoTapLim],
                     max_iter=int(dict_config[const.strPFMaxIter]))
                t7 = time.time()
                if pfconf.bMonitorAllAssets:
                    Lines = pf_object.app.GetCalcRelevantObjects('*.ElmLne')
                    lLines, lTrafos = ([] for i in range(2))
                    for line in Lines:
                        lLines.append(line.loc_name)
                    Trafos = pf_object.app.GetCalcRelevantObjects('*.ElmTr2')
                    for Trafo in Trafos:
                        lTrafos.append(Trafo.loc_name)
                else:
                    lLines, lTrafos = sio.readLineTrafoNames(assets=dfAssetData)
                dictLines, dictTrafos = sio.readLineTrafoNameThreshold(assets=dfAssetData)
                output_path = strScenario_folder + "\\"
                output_folder = output_path + "DOWNSCALING_FACTOR" + "\\"
                t8 = time.time()

                if bEvents:
                    print('Reading EVENTS file: ' + eventsFile)
                    dfAllEvents = fv.readEventFile(eventsFile, const.strDataSheet)
                    if 'SCENARIO' in dfAllEvents.columns:
                        dfEvents = dfAllEvents[dfAllEvents['SCENARIO']==scenario]
                        if dfEvents.empty:
                            bEvents = False
                    else:
                        dfEvents = dfAllEvents
                else:
                    dfEvents = pd.DataFrame()

                if bSwitch:
                    print('Reading SWITCH file: ' + switchFile)
                    dfAllSwitch = fv.readSwitchFile(switchFile, const.strDataSheet)
                    if 'SCENARIO' in dfAllSwitch.columns:
                        dfSwitch = dfAllSwitch[dfAllSwitch['SCENARIO']==scenario]
                        if dfSwitch.empty:
                            bSwitch = False
                    else:
                        dfSwitch = dfAllSwitch
                else:
                    dfSwitch = pd.DataFrame()

                pb=20
                pbar.updateProgressBar(pb, "Preparing Load Flow Calculations")

                status, msg, dictLoadingLines, dictLoadingTrafos, dfDownScalingFactor_grid, dfAllTrafos_grid_PQnom, \
                dictLinePowerBus1, dictLineIBus1, dictLineIBus2, dictTrafoPowerLv, dictTrafoIBus1, dictTrafoIBus2= \
                    pffun.PF_CalcAssetLoading(isAuto=isAuto, PSArunID=PSArunID, bEvents=bEvents, dfEvents=dfEvents,
                     bSwitch=bSwitch, dfSwitch=dfSwitch, plout=plout,unplout=unplout,
                     target_NeRDA_folder=NeRDA_output_path,pf=pf_object,no_days=number_of_days,
                     no_half_hours=number_of_half_hours,lMonitoredLineAssets=lLines,lMonitoredTrafoAssets=lTrafos,
                     lOutageAssets=lOutageAssets,lOutageAssetTypes=lOutageAssetTypes,lDfOutages=lDfOutages,
                     output_folder=output_folder,power_factor=power_factor,lagging=lagging,start_time_row=start_time_row,
                     bDebug=bDebug,dir_lnelding=output_path,dir_trafolding=output_path,PSA_running_mode=PSA_running_mode,
                     dictSIAFeeder=dictSIAFeeder,dictSIAGen=dictSIAGen, lNamesGrid=lNamesGrid, dfDownScalingFactor_grid=dfDownScalingFactor_grid,
                     dfAllTrafos_grid_PQnom=dfAllTrafos_grid_PQnom, PSAfolder=filepath, 
                     number_of_days=number_of_days, number_of_half_hours=number_of_half_hours)
                t9 = time.time()

                if status == const.PSAok:

                    pb=94
                    pbar.updateProgressBar(pb, "Processing Line Overloading Results")

                    status_msg_lne,lCombinedLineEvents,lCombinedLineConstraints,bLineEvents = pffun.PF_CalcOverLoading(pf=pf_object, strHHmode=dict_config[const.strHHMode], type="line", dictLoadingAssets=dictLoadingLines,
                                                              threshold=lne_thresh, start_date=start_date,PSA_running_mode=PSA_running_mode, dfAsset=dfAssetData,
                                                              dir_line_loading_data=output_path + "LN_LDG_DATA" + "\\",
                                                              dir_trafo_loading_data=output_path + "TX_LDG_DATA" + "\\",
                                                              dictLinePowerBus1=dictLinePowerBus1, dictLineIBus1=dictLineIBus1, dictLineIBus2=dictLineIBus2,
                                                              dictTrafoPowerLv=dictTrafoPowerLv, dictTrafoIBus1=dictTrafoIBus1, dictTrafoIBus2=dictTrafoIBus2,
                                                              number_of_days=number_of_days, number_of_half_hours=number_of_half_hours)
                    t10 = time.time()

                    pb=96
                    pbar.updateProgressBar(pb, "Processing Transformer Overloading Results")

                    status_msg_trafo,lCombinedTrafoEvents,lCombinedTrafoConstraints,bTrafoEvents = pffun.PF_CalcOverLoading(pf=pf_object, strHHmode=dict_config[const.strHHMode],type="trafo", dictLoadingAssets=dictLoadingTrafos,
                                                            threshold=trafo_thresh, start_date=start_date,PSA_running_mode=PSA_running_mode, dfAsset=dfAssetData,
                                                            dir_line_loading_data=output_path + "LN_LDG_DATA" + "\\",
                                                            dir_trafo_loading_data=output_path + "TX_LDG_DATA" + "\\",
                                                            dictLinePowerBus1=dictLinePowerBus1,dictLineIBus1=dictLineIBus1,
                                                            dictLineIBus2=dictLineIBus2,dictTrafoPowerLv=dictTrafoPowerLv,
                                                            dictTrafoIBus1=dictTrafoIBus1,dictTrafoIBus2=dictTrafoIBus2,
                                                            number_of_days=number_of_days, number_of_half_hours=number_of_half_hours)
                    t11 = time.time()

                    pb=98
                    pbar.updateProgressBar(pb, "Saving Flex Requirements Results")

                    msg = status_msg_lne + "\n" + status_msg_trafo
                    dfCombinedEvents = pffun.PF_CombineEvents(bLineEvents=bLineEvents,bTrafoEvents=bTrafoEvents,
                                    lCombinedLineEvents=lCombinedLineEvents,lCombinedTrafoEvents=lCombinedTrafoEvents)
                    dfCombinedEvents['req_id'] = [i for i in range(len(dfCombinedEvents))]
                    dfCombinedEvents.set_index('req_id')

                    dfCombinedConstraints = pffun.PF_CombineConstraints(bLineEvents=bLineEvents, bTrafoEvents=bTrafoEvents,
                                      lCombinedLineConstraints=lCombinedLineConstraints,lCombinedTrafoConstraints=lCombinedTrafoConstraints)
                    dfCombinedConstraints['req_id'] = [i for i in range(len(dfCombinedConstraints))]
                    dfCombinedConstraints.set_index('req_id')
                    t12 = time.time()

                    pffun.PF_OutputFlexReqts(bLineEvents, bTrafoEvents, dfCombinedEvents, dir_output=strScenario_folder
                                                                            + "\\" + "FLEX_REQTS", PSArunID=PSArunID)
                    pffun.PF_OutputConstraints(bLineEvents, bTrafoEvents, dfCombinedConstraints, dir_output=strScenario_folder
                                                                            + "\\" + "CONSTRAINTS",PSArunID=PSArunID)
                    
                    t13 = time.time()

                    #### THIS HAS ONLY OUTPUT FLEX AND CONSTRAINTS FILES TO SCENARIO LEVEL FOLDERS
                    #### OUTPUT TO PSA_RUN_ID AND SND SHARED FOLDERS AT NEXT STAGE 

                    end_time = datetime.now()
                    pf_interface.PF_exit(pf_object)
                    t14 = time.time()

                    with open(perf_file, 'a', newline='') as f:
                        writer = csv.writer(f)
                        writer.writerow(
                            [str(t1 - t0), str(t2 - t1), str(t3 - t2), str(t4 - t3), str(t5 - t4),
                            str(t6 - t5), str(t7 - t6), str(t8 - t7), str(t9 - t8), 
                            str(t10 - t9), str(t11 - t10), str(t12 - t11), str(t13 - t12), str(t14 - t13), datetime.now()])
                        f.close()

                    if not isAuto:
                        pb=100
                        pbar.updateProgressBar(pb, "Run Finished - Press button to Continue")
                    else:
                        pb=100
                        pbar.updateProgressBar(pb, "Run Finished")

                    print("\n\n################################################################")
                    print(f"SIMULATION FINISHED at {end_time}.")
                    print("################################################################\n\n")
    return status, msg, end_time, dictSIAFeeder, dictSIAGen, dictSIAGroup, start_time, NeRDA_output_path, \
           dfDownScalingFactor_grid,dfAllTrafos_grid_PQnom, dfCombinedEvents, dfCombinedConstraints
