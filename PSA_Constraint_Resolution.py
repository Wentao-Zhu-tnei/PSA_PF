import os
import logging
from os.path import exists
from datetime import datetime
import pandas as pd
import numpy as np
import PSA_Calc_Flex_Reqts as calcflex
import PF_Config as pfconf
import PSA_SND_Constants as const
import PSA_SND_Utilities as ut
import PSA_PF_SF as sf
import PSA_PF_Functions as pffun
import PSA_SIA_Data_Input_Output as sio
import PSA_NeRDA_Data_Input_Output as nrd
import powerfactory_interface as pf_interface
import PSA_Events_Input_Output as eventsio
import PSA_Switch_Input_Output as switchio
import PSA_File_Validation as fv
import UsefulPandas
import shutil

def createLogger(PSAfolder):
    # Create a log file in the PSAfolder with the current date in the name
    log_filename = os.path.join(PSAfolder, "CRlog.log")

    # Create a file handler that writes to the log file in append mode
    file_handler = logging.FileHandler(log_filename, mode="a")

    # Set the log message format
    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    file_handler.setFormatter(formatter)

    # Create a logger with the file handler
    logger = logging.getLogger(__name__)
    logger.addHandler(file_handler)
    logger.setLevel(logging.DEBUG)

    return logger

def getFilePSArunID(fname):
    PSArunID = "XXX_YYYY-MM-DD-hh-mm"
    lp = len(PSArunID)
    fpath = "C:\\Failed\\"
    ln1 = len(fname)

    if fname.find(const.strSNDResponses) >= 0:
        s2 = fname.find(const.strSNDResponses)
        ln2 = len(const.strSNDResponses)
    elif fname.find(const.strSNDContracts) >= 0:
        s2 = fname.find(const.strSNDContracts)
        ln2 = len(const.strSNDContracts)
    elif fname.find(const.strSNDCandidateResponses) >= 0:
        s2 = fname.find(const.strSNDCandidateResponses)
        ln2 = len(const.strSNDCandidateResponses)
    elif fname.find(const.strSNDCandidateContracts) >= 0:
        s2 = fname.find(const.strSNDCandidateContracts)
        ln2 = len(const.strSNDCandidateContracts)
    else:
        return PSArunID, fpath

    fpath = fname[:-(lp + ln2 + 5)]
    ln3 = len(fpath)
    fn = fname[ln3:]
    PSArunID = fn[:-(ln2 + 5)]

    return PSArunID, fpath


def getPSArunIDdate(PSArunID):
    strDate = PSArunID[4:-6]
    return strDate


def getPSArunIDtime(PSArunID):
    strTime = PSArunID[15:]
    return strTime

def PSA_constraint_resolution(bOutput2SND, bAuto, SNDfile):
    status = const.PSAok
    msg = os.path.basename(SNDfile)
    print("*****************************************************")
    print("PSA_Contraint_Resolution: START")
    print("Called with: " + msg)
    print("*****************************************************")

    dict_config = ut.PSA_SND_read_config(const.strConfigFile)
    root_working_folder = dict_config[const.strWorkingFolder]
    ldfReqid = []
    # get PSArunID from filename
    vn, PSArunID, PSAfolder = sf.getFilePSArunID(SNDfile)

    logger = createLogger(PSAfolder)
    resolution_start = datetime.now()
    logger.info(f"{PSArunID}-{vn}: Constraint resolution process started at {resolution_start}")  # Log the start time of the constraint resolution process

    ### LOAD RUN TIME PARAMETERS ###
    inFile = PSAfolder + PSArunID + const.strRunTimeParams
    if exists(inFile):
        dict_params = ut.readPSAparams(inFile)
    else:
        status = const.PSAfileExistError
        msg = "Run time Params file doesn't exist: " + inFile
        return status, msg

    dict_config = ut.PSA_SND_read_config(const.strConfigFile)

    # if in Auto mode being run by S&D process and assume multiprocess
    if bAuto:
        multiProcess = True
    else:
        multiProcess = False
    pf_object, status, msg = pf_interface.run(dict_config[const.strPFUser], multiProcess)

    if status != const.PSAok:
        msg = "PSA_Constraint_Resolution: " + msg
        return status, msg
    pf_interface.activate_project(pf_object, dict_params[const.strBSPModel],
                                    os.path.join(str(PSAfolder),str(const.strINfolder).replace("\\","")))
    
    if (not bAuto) and pfconf.bCreateRef:
        outputFile = os.path.join(dict_config[const.strWorkingFolder], const.strRefAssetFile)
        pffun.PF_getMasterRefList(pf_object, outputFile)
        print("Reference list of PowerFactory Assets has been updated!")

    # validate asset names
    lRefLine, lRefTrafo, lRefSyncMachine, lRefSwitch, lRefCoupler = fv.readRefAssetNames(
            os.path.join(dict_config[const.strWorkingFolder],const.strRefAssetFile))
    status, msg = fv.validateSNDFileNames(lRefSyncMachine, SNDfile)
    
    if status != const.PSAok:
        return status, msg
    start_date = getPSArunIDdate(PSArunID)
    start_time = getPSArunIDtime(PSArunID)

    if SNDfile.find(const.strSNDCandidateResponses) >= 0:
        bResponse = True
    elif SNDfile.find(const.strSNDCandidateContracts) >= 0:
        bResponse = False
    else:
        status = const.PSAfileTypeError
        msg = "Unknown file type: " + SNDfile
        return status, msg

    ### NAME OF OUTPUT FILE
    if bResponse:
        outFile1 = PSAfolder + PSArunID + "_PSA_CANDIDATE-RESPONSES" + "-" + vn + ".xlsx"
        outFile2 = dict_config[const.strFileDetectorFolder] + "//" + PSArunID + "_PSA_CANDIDATE-RESPONSES" + "-" + vn + ".xlsx"
    else:
        outFile1 = PSAfolder + PSArunID + "_PSA_CANDIDATE-CONTRACTS" + "-" + vn + ".xlsx"
        outFile2 = dict_config[const.strFileDetectorFolder] + "//" + PSArunID + "_PSA_CANDIDATE-CONTRACTS" + "-" + vn + ".xlsx"

    ### TEMP FILE FOLDER CREATION ###
    tempFolder = PSAfolder + "//" + const.strTempCRfiles
    if not os.path.exists(tempFolder):
        os.mkdir(tempFolder)
    tempFolder = tempFolder + "//" + vn + "//"
    if not os.path.exists(tempFolder):
        os.mkdir(tempFolder)

    ### EXTRACT ORIGINAL USER INPUT PARAMETERS ###
    bDebug = pfconf.bDebug
    number_of_days = int(dict_params[const.strDays])
    number_of_half_hours = int(dict_params[const.strHalfHours])
    power_factor = float(dict_params[const.strPowFctr])
    lagging = (str(dict_params[const.strPowFctrLag]) == "True")
    file_detector_folder = dict_config[const.strFileDetectorFolder]

    ### LOAD SND INPUT FILE ###
    dfSNDoutput = pd.read_excel(SNDfile)
    dfSNDoutput = ut.aggregateSNDresponses(dfSNDoutput)
    dfSNDoutput.to_excel(tempFolder+"\\"+PSArunID+const.strSNDCandidateResponses+"-DEDUP"+"-"+vn+".xlsx",index=False)
    calculate_historical = dfSNDoutput['calculate_historical'][0]
    input_folder = PSAfolder + const.strINfolder
    pf_proj_dir = input_folder

    ### LOAD ORIGINAL CONSTRAINTS FILE ###
    PSA_const_file = PSAfolder + PSArunID + "_PSA_CONSTRAINTS.xlsx"
    dfPSAconstraints = pd.read_excel(PSA_const_file)
    dfPSAconstraints['req_id'] = dfPSAconstraints['req_id'].astype(str)

    ### LOAD ASSET DATA FILE ###
    inFile = os.path.join(input_folder, dict_params[const.strAssetData])
    if exists(inFile):
        dfAssetData = pd.read_excel(os.path.join(input_folder, inFile), sheet_name=const.strAssetSheet)
        dfSIAData = pd.read_excel(os.path.join(input_folder, inFile), sheet_name=const.strSIASheet)
        dfNeRDAData = pd.read_excel(os.path.join(input_folder, inFile), sheet_name=const.strNeRDASheet)
    else:
        status = const.PSAfileExistError
        msg = "Asset Data File doesn't exist: " + inFile
        return status, msg

    start_date_time = start_date + " " + start_time.split("-")[0] + ":" + start_time.split("-")[1] + ":00"
    PSA_running_mode = sf.PSA_PF_runningmode(dict_params)

    ### LOAD SIA DATA FROM FILE OR API ###
    SIA_start = datetime.now()  # Log the start time of loading SIA data
    logger.debug(f"Started loading SIA for {vn} at {SIA_start}")
    SIAfolder = os.path.join(PSAfolder, "1 - SIA_DATA")
    SIAoutputfolder = os.path.join(PSAfolder, "1 - SIA_DATA_NEW")

    if not os.path.isdir(SIAoutputfolder):
        os.mkdir(SIAoutputfolder)
    if calculate_historical:
        print("\n################################################################")
        print(f"SIA - Using historical data from folder: [{SIAfolder}]")
        print("################################################################\n")
        dictSIAFeeder, dictSIAGen, dictSIAGroup = sf.PSA_PF_readSIA(SIAfolder)
        SIA_start_date_time = start_date_time
    else:
        token, status, msg = sio.SIA_get_token(dict_config[const.strSIAUsr], dict_config[const.strSIAPwd],
                                               int(dict_config[const.strSIATimeOut]))
        if status != const.PSAok:
            return status, msg
        else:
            SIA_start_time, SIA_start_date_time, status, msg, dictSIAFeeder, dictSIAGen, dictSIAGroup = sio.loadSIAData(
                strbUTC=dict_config[const.bUTC], assets=dfSIAData, timeout=int(dict_config[const.strSIATimeOut]),
                OutputFilePath=SIAoutputfolder, token=token,bDebug=bDebug, PSArunID=PSArunID,
                number_of_days=number_of_days, number_of_half_hours=number_of_half_hours)
            if status != const.PSAok:
                return status, msg
            else:
                start_date = str(start_time).split(' ')[0]

    SIA_end = datetime.now()   # Log the end time of loading SIA data
    logger.debug(f"Finished loading SIA data at {SIA_end}. Time taken: {SIA_end - SIA_start}")

    ### LOAD NERDA DATA FROM FILE OR API ###
    NeRDA_start = datetime.now()  # Log the start time of loading SIA data
    logger.debug(f"Started loading NeRDA for {vn} at {NeRDA_start}")

    NeRDAfolder = os.path.join(PSAfolder, "2 - NeRDA_DATA")
    if calculate_historical:
        print("\n################################################################")
        print(f"NeRDA - Using historical data from folder: [{NeRDAfolder}]")
        print("################################################################\n")
        inFile = NeRDAfolder + "//" + PSArunID + "-NeRDA_DATA.xlsx"
        
        if not os.path.isfile(inFile):
            copyNeRDAfile1 = dict_config[const.strWorkingFolder] + "\\" + const.strDataResults + "\\AUT_2023-MM-DD-HH-MM-NeRDA_DATA.xlsx"        
            copyNeRDAfile2 = NeRDAfolder + "\\AUT_2023-MM-DD-HH-MM-NeRDA_DATA.xlsx"
            shutil.copyfile(copyNeRDAfile1, copyNeRDAfile2)
            inFile = copyNeRDAfile2

        dfNeRDA = pd.read_excel(inFile)
    else:
        if pfconf.bUseNeRDA:
            print("NeRDA - Loading NeRDA data from API: " + NeRDAfolder)
            status, msg, dfNeRDA = nrd.loadNeRDAData(assets=dfNeRDAData, output_path=NeRDAfolder + "\\" + PSArunID,
                                        user=dict_config[const.strNERDAUsr], pwd=dict_config[const.strNERDAPwd])
        if status != const.PSAok:
            return status, msg

    if pfconf.bUseNeRDA:
        print("=" * 65 + "\n" + "Loading NeRDA data" + "\n" + "=" * 65)
        dfNeRDACopy = dfNeRDA.copy()
        dfNeRDACopy = calcflex.PSA_filter_NeRDA(dfNeRDACopy, dfNeRDAData)
        dfElmCoup_ALL = pf_object.get_all_circuit_breakers_coup()
        dfStaSwitch_ALL = pf_object.get_all_circuit_breakers_switch()
        pf_object.enableLocalCache()
        calcflex.PSA_update_switch(pf_object, dfElmCoup_ALL, dfStaSwitch_ALL, dfNeRDACopy)
        pf_object.disableLocalCache()
        pf_object.updatePFdb()

    NeRDA_end = datetime.now()  # Log the end time of loading NeRDA data
    logger.debug(f"Finished loading NeRDA data at {NeRDA_end}. Time taken: {NeRDA_end - NeRDA_start}")

    # If mode is AUT then BASE must exist so use that folder
    if dict_params[const.strRunMode] == const.strRunModeAUT:
        strDownscalingFactorFolder = PSAfolder + const.strBfolder + "\\DOWNSCALING_FACTOR\\"
    # Otherwise for MAN and SEMI-AUT check MAINT and CONT
    elif dict_params[const.strMaint] == "False" and dict_params[const.strCont] == "False":
        strDownscalingFactorFolder = PSAfolder + const.strBfolder + "\\DOWNSCALING_FACTOR\\"
    elif dict_params[const.strMaint] == "True" and dict_params[const.strCont] == "False":
        strDownscalingFactorFolder = PSAfolder + const.strMfolder + "\\DOWNSCALING_FACTOR\\"
    elif dict_params[const.strMaint] == "False" and dict_params[const.strCont] == "True":
        strDownscalingFactorFolder = PSAfolder + const.strCfolder + "\\DOWNSCALING_FACTOR\\"
    elif dict_params[const.strMaint] == "True" and dict_params[const.strCont] == "True":
        strDownscalingFactorFolder = PSAfolder + const.strMCfolder + "\\DOWNSCALING_FACTOR\\"

    file_found = False
    dfDownScalingFactor_grid = pd.DataFrame()
    for r, d, f in os.walk(strDownscalingFactorFolder):
        for file in f:
            if "PSA_DOWNSCALING_FACTORS_PQNOM" in file:
                print(file)
                print(f'Read xlsx file as dataframe')
                dfDownScalingFactor_grid = pd.read_excel(
                    os.path.join(strDownscalingFactorFolder, "ALL_PSA_DOWNSCALING_FACTORS_PQNOM.xlsx"))
                file_found = True

    if file_found == False:
        status = const.PSAfileReadError
        msg = f"Donwscaling Factor file not found in {strDownscalingFactorFolder}"
        return status, msg
    else:
        dfDownScalingFactor_grid['Name_grid'] = dfDownScalingFactor_grid['Name'] + '_' + dfDownScalingFactor_grid['Grid']

    print("=" * 65)
    print("Input files read")
    print("PSArunID = " + PSArunID)
    print("SIA start_date_time = " + SIA_start_date_time)
    print("PSA start_date_time = " + start_date_time)
    print("PSA running mode = " + PSA_running_mode)
    print("BSP Model = " + dict_params[const.strBSPModel])
    print("Debug = " + str(bDebug))
    print("Power Factor = " + str(power_factor))
    print("Lagging = " + str(lagging))
    print("SIA use historical = " + str(calculate_historical))
    print("Output to SND = " + str(bOutput2SND))
    print("Automatic = " + str(bAuto))

    ### get list of lines and transformers ###
    all_loads_df = pf_object.get_loads_names_grid_obj()

    ### CREATE EMPTY OUTPUT DATEFRAME ###
    dfOutput = pd.DataFrame(columns=["req_id", "bsp", "primary", "feeder", "secondary", "terminal", "constrained_pf_id",
                                     "busbar_from", "busbar_to", "scenario", "start_time", "duration",
                                     "residual_req_power_kw", "loading_pct_initial", "loading_pct_final",
                                     "loading_threshold_pct", "tot_offered_power_kw", "tot_responses"])
    verbose = True

    groupedbyscenario = dfSNDoutput.groupby(['scenario'])
    lScenarios = list(set(dfSNDoutput['scenario']))
    network_element_df_all_steps_all_scenarios = {}
    #### COUNTER FOR OUTPUT FILE ####
    iReq = 0

    if dict_params[const.strbEvent] == "True":
        bEvent = True
    else:
        bEvent = False

    if dict_params[const.strSwitch] == "True":
        bSwitch = True
    else:
        bSwitch = False

    if bEvent == True:
        dfAllEvents = fv.readEventFile(PSAfolder + const.strINfolder + dict_params[const.strEventFile], const.strDataSheet)
        print("-" * 20)
        print(f"Events file {dict_params[const.strEventFile]} has been read.")
        print("-" * 20)
    if bSwitch:
        dfAllSwitch = fv.readSwitchFile(os.path.join(PSAfolder+const.strINfolder+dict_params[const.strSwitchFile]),
                                     const.strDataSheet)
        print("-" * 20)
        print(f"Switch file {dict_params[const.strSwitchFile]} has been read.")
        print("-" * 20)


    #### MAIN LOOP ####
    pf_object.enableLocalCache()
    for idx, scenario in enumerate(lScenarios):
        print(
            f"\n\n================================\n **** {idx + 1}/{len(lScenarios)} Scenario: [{scenario}]****\n================================")
        if bEvent:
            if 'SCENARIO' in dfAllEvents.columns:
                dfEvents = dfAllEvents[dfAllEvents['SCENARIO'] == scenario]
                if dfEvents.empty:
                    bEvent = False
            else:
                dfEvents = dfAllEvents
            dictDfServices = eventsio.mapServices(strbUTC=dict_config[const.bUTC],
                                                  start_time=PSArunID[4:14] + " " + "00:00:00",
                                                  dfEvents=dfEvents, PSAfolder=PSAfolder,
                                                  number_of_days=number_of_days,
                                                  number_of_half_hours=number_of_half_hours)
            print("Events have been mapped.")
            print("-" * 20)

        if bSwitch:
            if 'SCENARIO' in dfAllSwitch.columns:
                dfSwitch = dfAllSwitch[dfAllSwitch['SCENARIO'] == scenario]
                if dfSwitch.empty:
                    bSwitch = False
            else:
                dfSwitch = dfAllSwitch
            dictSwitchStatus = switchio.mapSwitchStatus(strbUTC=dict_config[const.bUTC],
                                                        start_time=PSArunID[4:14] + " " + "00:00:00",
                                                        dfSwitch=dfSwitch, PSAfolder=PSAfolder,
                                                        number_of_days=number_of_days,
                                                        number_of_half_hours=number_of_half_hours)
            print("Switch statuses have been mapped.")
            print("-" * 20)

        dfScenario = groupedbyscenario.get_group(scenario)

        # Read in the component status array based on scenario name
        if scenario == "MAINT" or scenario == "CONT" or scenario == "MAINT_CONT":
            dictDfTrafos_2w, dictDfLine, dictDfSynm, dictDfLoad = ut.readOutageFile(scenario, PSAfolder)

        ### Groupby start_time ###
        groupedbystarttime = dfScenario.groupby(['start_time'])
        lStarttime = list(set(dfScenario['start_time']))
        if verbose == True:
            print("================================================")
            print(dfScenario.shape)
            print("lStarttime are:", lStarttime)
            print("================================================")

        ### Iterate through Constraint Starttime ###
        for idx, starttime in enumerate(lStarttime):
            time_diff = (datetime.strptime(starttime, '%Y-%m-%d %H:%M:%S') -
                         datetime.strptime(SIA_start_date_time, '%Y-%m-%d %H:%M:%S')).days
            if time_diff < 0:
                ispast = True
            else:
                ispast = False
                print(f"Constraint happened at {starttime} is in the past; it is skipped.")
            if ispast == False:
                print(f"================================\n \t{idx + 1}/{len(lStarttime)} Constrained time: [{starttime}]")
                # Update componenet status according to [scenario] and [starttime]

                dfStarttime = groupedbystarttime.get_group(starttime)
                ### Groupby req_ids ###
                groupedbyreqid = dfSNDoutput.groupby(['req_id'])
                setReqids = list(set(dfStarttime['req_id']))
                if verbose == True:
                    print("================================================")
                    print(dfStarttime.shape)
                    print("\tsetReqids are:", setReqids)
                    print("================================================")

                # psa_start_datetime = start_date_time
                status, i_half_hour, i_day = sf.PSA_PF_find_datetime_idx(start_date_time,starttime,
                                                                         verbose=False)
                i_day_shifted = i_day - (datetime.strptime(SIA_start_date_time, '%Y-%m-%d %H:%M:%S').replace(hour=0, minute=0, second=0, microsecond=0) -
                         datetime.strptime(start_date_time, '%Y-%m-%d %H:%M:%S').replace(hour=0, minute=0, second=0, microsecond=0)).days
                if bEvent:
                    eventsio.updateServices(pf_object, dictDfServices, i_half_hour, i_day, verbose=True)
                if bSwitch:
                    switchio.updateSwitchStatus(pf_object, dictSwitchStatus, i_half_hour, i_day, verbose=True)
                    print("-" * 20)
                    print(f"Switch status is updated for day {i_day} and half hour {i_half_hour}.")
                    print("-" * 20)
                if scenario == "MAINT" or scenario == "CONT" or scenario == "MAINT_CONT":
                    pffun.PF_update_outages(pf_object, dictDfLine, dictDfTrafos_2w, dictDfSynm, dictDfLoad, i_half_hour, i_day, verbose=True)
                    print(f'\t --Done updating outages for {scenario} {starttime}')
                if status == const.PSAok:
                    ## Load SIA gen data for non- _F assets that have SIA data ##
                    for Genname, dictSIAGenParams in dictSIAGen.items():
                        Genname = Genname.replace(".xlsx", "")
                        oGen = pf_object.app.GetCalcRelevantObjects(Genname + '.ElmSym')[0]
                        oGen.pgini = dictSIAGenParams['GEN_MW'].iloc[i_half_hour][i_day_shifted]
                    print('\t --Done updating Generator setpoints according to SIA')

                all_loads_df = sf.downscale_SIA_to_loads(all_loads_df, dictSIAFeeder, i_day_shifted, i_half_hour,
                                                         dfDownScalingFactor_grid, power_factor, lagging, verbose=False)

                print('\t --Done downscaling P/Q setpoints of loads based on SIA')

                ### Iterate through req_ids ###
                for idx, reqid in enumerate(setReqids):
                    reqloop_start = datetime.now()  # Log the start time of main loop
                    logger.debug(f"{PSArunID}-{vn}: Started solving constraints for {reqid} at {reqloop_start}")

                    print(f"\t\t{idx + 1}/{len(setReqids)} - Requirement ID [{reqid}] - day:[{i_day}] hh:[{i_half_hour}] ")
                    dfReqid = groupedbyreqid.get_group(reqid)
                    constrained_pf_id = dfReqid.iloc[0]['constrained_pf_id']
                    constrained_pf_type = dfReqid.iloc[0]['constrained_pf_type']

                    #### SET OUTPUT FILE VALUES ###'
                    dfOutput._set_value(index=iReq, col='req_id', value=dfReqid.iloc[0]['req_id'])
                    dfOutput._set_value(index=iReq, col='bsp', value=dfReqid.iloc[0]['bsp'])
                    dfOutput._set_value(index=iReq, col='primary', value=dfReqid.iloc[0]['primary'])
                    dfOutput._set_value(index=iReq, col='feeder', value=dfReqid.iloc[0]['feeder'])
                    dfOutput._set_value(index=iReq, col='secondary', value=dfReqid.iloc[0]['secondary'])
                    dfOutput._set_value(index=iReq, col='terminal', value=dfReqid.iloc[0]['terminal'])
                    dfOutput._set_value(index=iReq, col='constrained_pf_id', value=dfReqid.iloc[0]['constrained_pf_id'])
                    dfOutput._set_value(index=iReq, col='busbar_from', value=dfReqid.iloc[0]['busbar_from'])
                    dfOutput._set_value(index=iReq, col='busbar_to', value=dfReqid.iloc[0]['busbar_to'])
                    dfOutput._set_value(index=iReq, col='scenario', value=dfReqid.iloc[0]['scenario'])
                    dfOutput._set_value(index=iReq, col='start_time', value=dfReqid.iloc[0]['start_time'])
                    dfOutput._set_value(index=iReq, col='duration', value=dfReqid.iloc[0]['duration'])
                    # dfOutput._set_value(index=iReq, col='residual_req_power_kw', value=999.999)


                    ### GET THE ID DATA FROM ORIGINAL CONSTRAINTS DATA ###
                    dfReqid_PSA = UsefulPandas.filter_df_single(dfPSAconstraints, 'req_id', dfReqid['req_id'].values[0],
                                                                regex=False)

                    loading_threshold_pct = dfReqid_PSA['loading_threshold_pct'].values[0]
                    flex_pf_id = dfReqid['flex_pf_id'].values[0]

                    ### Groupby resp_ids ###
                    if bResponse:
                        groupedbyresponse = dfReqid.groupby(['resp_id'])
                    else:
                        groupedbyresponse = dfReqid.groupby(['con_id'])
                    if bResponse:
                        setRespids = list(set(dfReqid['resp_id']))
                    else:
                        setRespids = list(set(dfReqid['con_id']))
                    if bResponse:
                        req_responses_all_agg = dfReqid.groupby('req_id').agg({'resp_id': 'count',
                                                                               'offered_power_kw': np.sum})
                    else:
                        req_responses_all_agg = dfReqid.groupby('req_id').agg({'con_id': 'count',
                                                                               'offered_power_kw': np.sum})
                    total_offered_power_kw = req_responses_all_agg['offered_power_kw'].values[0]
                    if bResponse:
                        total_num_responses = req_responses_all_agg['resp_id'].values[0]
                    else:
                        total_num_responses = req_responses_all_agg['con_id'].values[0]
                    # dfOutput._set_value(index=iReq, col='loading_pct', value=dfReqid_PSA['loading_pct'].values[0])
                    dfOutput._set_value(index=iReq, col='loading_threshold_pct', value=loading_threshold_pct)
                    dfOutput._set_value(index=iReq, col='tot_responses', value=total_num_responses)
                    print(f'\t\t total_offered_power_kw: {total_offered_power_kw},total_responses: {total_num_responses}')

                    if verbose == True:
                        print("================================================")
                        print(dfReqid.shape)
                        if bResponse:
                            print("\t\tsetRespids are:", setRespids)
                        else:
                            print("\t\tsetConids are:", setRespids)
                        print("================================================")

                    ### Prepare load flow ###
                    pf_object.prepare_loadflow(ldf_mode=dict_config[const.strPFLFMode],
                                               algorithm=dict_config[const.strPFAlgrthm],
                                               trafo_tap=dict_config[const.strPFTrafoTap],
                                               shunt_tap=dict_config[const.strPFShuntTap],
                                               zip_load=dict_config[const.strPFZipLoad],
                                               q_limits=dict_config[const.strPFQLim],
                                               phase_shift=dict_config[const.strPFPhaseShift],
                                               trafo_tap_limit=dict_config[const.strPFTrafoTapLim],
                                               max_iter=int(dict_config[const.strPFMaxIter]))

                    ## Run Load Flow For Each Req_ID (single load flow to get an initial set of conditions)##
                    print('=' * 65)
                    pf_object.updatePFdb()
                    print(f"check for power flow for Requirement ID [{reqid}] - day:[{i_day}] hh:[{i_half_hour}]")
                    ierr1 = pf_object.run_loadflow()
                    if ierr1 == 0:
                        status = const.PSAok
                        print(
                            f"\t\t\tLoad Flow SUCCESSFUL for Initial Conditions\n \t\t\t(before flex) - Day: {i_day}, Half Hour:{i_half_hour}")
                    else:
                        status = const.PSAPFError
                        print(
                            f"\t\t\tLoad flow ERROR for Initial Conditions\n \t\t\t(before flex) - Day: {i_day},  Half Hour:{i_half_hour}")
                    print('=' * 65)

                    if status == const.PSAok:
                        network_element_df_all_steps = pd.DataFrame(columns=['PF_Object', 'Name',
                                                                             'Type', 'Grid',
                                                                             'From_HV', 'To_LV',
                                                                             'From_HV_Status', 'To_HV_Status',
                                                                             'P_HV', 'Q_HV', 'S_HV', 'I_HV', 'V_HV',
                                                                             'P_LV', 'Q_LV', 'S_LV', 'I_LV', 'V_LV',
                                                                             'I_nom_act_kA_lv', 'I_nom_act_kA_hv',
                                                                             'I_max_act',
                                                                             'loading_threshold_pct',
                                                                             'Loading', 'loading_diff_perc',
                                                                             'Ploss', 'Qloss', 'Overload',
                                                                             'OutService', 'CIM_ID', 'Service',
                                                                             'P_max_id', 'I_exceed', 'I_exceed_val', 'P_req',
                                                                             ])
                        network_element_df_res_flex = PSA_residual_flex_only(status, pf_object,
                                                                             constrained_pf_id, constrained_pf_type,
                                                                             loading_threshold_pct)
                        network_element_df_all_steps.loc['initial'] = network_element_df_res_flex.values[0]

                    ### Iterate through Responses ###
                    dfAggResp = pd.DataFrame(columns=['Flex Asset', 'Pgini'])
                    for idx, respid in enumerate(setRespids):
                        print('=' * 65)
                        if bResponse:
                            print(f"\t\t\t{idx + 1}/{len(setRespids)} - Response ID [{respid}] ")
                        else:
                            print(f"\t\t\t{idx + 1}/{len(setRespids)} - Contract ID [{respid}] ")
                        dfRespid = groupedbyresponse.get_group(respid).iloc[0]
                        if bResponse:
                            resp_id = dfRespid['resp_id']
                        else:
                            resp_id = dfRespid['con_id']
                        print(f"================================\n\t\t\t resp_id {resp_id}\n")
                        flex_pf_id = dfRespid['flex_pf_id']
                        offered_power_mw = dfRespid['offered_power_kw'] / 1e3  # this offered power has to be converted to MW for PowerFactory

                        if bResponse:
                            print(
                                f'\t\t\t{idx + 1}/{len(setRespids)} Req ID {reqid} - Resp ID: {resp_id} - \n\t\t\tNetwork element: {constrained_pf_id} - Flex asset: {flex_pf_id} ')
                        else:
                            print(
                                f'\t\t\t{idx + 1}/{len(setRespids)} Req ID {reqid} - Contract ID: {resp_id} - \n\t\t\tNetwork element: {constrained_pf_id} - Flex asset: {flex_pf_id} ')

                        ## Change Flex Asset Dispatch For Each Resp_ID ##
                        oGen_flex = pf_object.app.GetCalcRelevantObjects(flex_pf_id + '.ElmSym')[0]
                        oGen_flex.pgini = offered_power_mw  # Update the dispatch power of the flex asset

                        lDataRow = {'Flex Asset':flex_pf_id, 'Pgini':oGen_flex.pgini}
                        dfAggResp = dfAggResp.append(lDataRow, ignore_index=True)

                        print(f'\t\t\tUpdated flex asset [{oGen_flex.loc_name}] Pgini to [{offered_power_mw}]MW')

                        if verbose == True:
                            print("================================================")
                            print(dfRespid.shape)
                            print("\t\t\tsetRespids are:", setRespids)
                            print("================================================")
                    dfAggResp.to_excel(tempFolder + f'AggregateResponses_{constrained_pf_id}_{reqid}.xlsx', index=False)
                    ######### After updating all Flex Assets that come in a bundle #########
                    ## Run load flow ##
                    print('=' * 65)
                    pf_object.updatePFdb()
                    ierr2 = pf_object.run_loadflow()
                    if ierr2 == 0:
                        status = const.PSAok
                        print(
                            f"\t\t\tLoad Flow SUCCESSFUL with Flex Asset [{flex_pf_id}] dispatched\n \t\t\t (after flex) - Day: {i_day}, Half Hour:{i_half_hour}")
                    else:
                        status = const.PSAPFError
                        print(
                            f"\t\t\tLoad flow ERROR with Flex Asset [{flex_pf_id}] dispatched\n \t\t\t (after flex) - Day: {i_day}, Half Hour:{i_half_hour}")
                    print('=' * 65)

                    if status == const.PSAok:
                        network_element_df_res_flex = PSA_residual_flex_only(status, pf_object,
                                                                             constrained_pf_id, constrained_pf_type,
                                                                             loading_threshold_pct)

                    network_element_df_all_steps.loc['final'] = network_element_df_res_flex.values[0]
                    network_element_df_all_steps['tot_responses'] = total_num_responses

                    # dfOutput._set_value(index=iReq, col='required_flex_kw',
                    #                     value=network_element_df_all_steps.loc['final']["res_flex_req_power_kw"])
                    # dfOutput._set_value(index=iReq, col='residual_req_power_kw',
                    #                     value=network_element_df_all_steps.loc['final']["res_flex_req_power_kw"])

                    dfOutput._set_value(index=iReq, col='loading_pct_initial',
                                        value=network_element_df_all_steps.loc['initial']["Loading"])
                    dfOutput._set_value(index=iReq, col='loading_pct_final',
                                        value=network_element_df_all_steps.loc['final']["Loading"])
                    dfOutput._set_value(index=iReq, col='residual_req_power_kw',
                                        value=network_element_df_all_steps.loc['final']["P_req"])
                    dfOutput._set_value(index=iReq, col='tot_offered_power_kw',
                                        value=total_offered_power_kw.item())
                    network_element_df_all_steps['tot_offered_power_kw'] = total_offered_power_kw
                    print("########################")
                    print(f"total offered power is {total_offered_power_kw}, the type of total_offered_power_kw is {type(total_offered_power_kw)}")
                    print("########################")
                    network_element_df_all_steps['reqID_day_hh'] = f'{reqid}_day_{i_day}_hh_{i_half_hour}'
                    network_element_df_all_steps_all_scenarios[
                        f'{reqid}_day_{i_day}_hh_{i_half_hour}'] = network_element_df_all_steps
                    print(
                        f'\t\t\tSaving intermediary file:\n ConRes_int_{flex_pf_id}_{constrained_pf_id}_{reqid}_{resp_id}.csv')

                    csvFile = tempFolder + f'ConRes_int_{constrained_pf_id}_{reqid}_{resp_id}.csv'
                    network_element_df_all_steps.to_csv(csvFile, index=True)

                    network_element_df_all_steps = None

                    # Log the end time of each critical operation and the time taken
                    reqloop_end = datetime.now()
                    logger.debug(f"{PSArunID}-{vn}: Ended solving requirement {reqid} at {reqloop_end}. Time taken: {reqloop_end - reqloop_start}")

                    ### NEXT OUTPUT FILE ROW ###
                    iReq = iReq + 1

                    ### Iterate through Responses to RESET values to zero ###
                    print("\t\t\tRestarting flex assets to zero")
                    for idx, respid in enumerate(setRespids):
                        print('=' * 65)
                        print(f"\t\t\t{idx + 1}/{len(setRespids)} - Response ID [{respid}] ")
                        dfRespid = groupedbyresponse.get_group(respid).iloc[0]

                        if bResponse:
                            resp_id = dfRespid['resp_id']
                        else:
                            resp_id = dfRespid['con_id']
                        flex_pf_id = dfRespid['flex_pf_id']
                        offered_power_mw = dfRespid['offered_power_kw'] / 1e3  # this offered power has to be converted to MW for PowerFactory

                        if bResponse:
                            print(
                                f'\t\t\t{idx + 1}/{len(setRespids)} Req ID {reqid} - Resp ID: {resp_id} - \n\t\t\tNetwork element: {constrained_pf_id} - Flex asset: {flex_pf_id} ')
                        else:
                            print(
                                f'\t\t\t{idx + 1}/{len(setRespids)} Req ID {reqid} - Contract ID: {resp_id} - \n\t\t\tNetwork element: {constrained_pf_id} - Flex asset: {flex_pf_id} ')

                        ## Change Flex Asset Dispatch For Each Resp_ID/Con_ID ##
                        oGen_flex = pf_object.app.GetCalcRelevantObjects(flex_pf_id + '.ElmSym')[0]
                        oGen_flex.pgini = 0  # Update the dispatch power of the flex asset
                        print(f'\t\t\tUpdated flex asset [{oGen_flex.loc_name}] Pgini to [{offered_power_mw}]MW')
                    ## Run load flow to make sure the changes take place##
                    print('=' * 65)
                    pf_object.updatePFdb()
                    ierr3 = pf_object.run_loadflow()
                    if ierr3 == 0:
                        status = const.PSAok
                        print(
                            f"\t\t\tLoad Flow SUCCESSFUL \n \t\t\t (after restart) - Day: {i_day}, Half Hour:{i_half_hour}")
                    else:
                        status = const.PSAPFError
                        print(f"\t\t\tLoad flow ERROR \n \t\t\t (after restart) - Day: {i_day}, Half Hour:{i_half_hour}")

                    print('=' * 65)
                    print("\t\t\tDONE restarting flex assets to zero")
                    print('"================================\n"================================\n')

    ### EXIT POWER FACTORY ###
    pf_object.disableLocalCache()
    pf_interface.PF_exit(pf_object)

    ### OUTPUT RESULTS ###
    print("OUTPUT FILE = " + outFile1)
    dfOutput.to_excel(outFile1, index=False)

    if bOutput2SND:
        print("SHARED OUTPUT FILE = " + outFile2)
        dfOutput.to_excel(outFile2, index=False)

    ### END OF FUNCTION ###
    print("*****************************************************")
    print("PSA_Contraint_Resolution: END")
    print("*****************************************************")

    # Log the end of the function
    resolution_end = datetime.now()
    logger.info(f"{PSArunID}-{vn}: Ended constrained resolution, total time taken: {resolution_end - resolution_start}")

    return status, msg


def PSA_residual_flex(status, pf_object, constrained_pf_id, loading_threshold_pct, flex_pf_id, iter_step,
                      network_element_df_all_steps):
    if status == const.PSAok:

        ## Network element parameters
        specific_network_element = constrained_pf_id + '*.*'
        specific_network_element = pf_object.app.GetCalcRelevantObjects(specific_network_element)[0]
        specific_network_element_type = specific_network_element.GetClassName()

        if specific_network_element_type == 'ElmLne':

            ## Get network element results from loadflow
            network_element_df = pf_object.get_line_terminal_result(constrained_pf_id, loading_threshold_pct)

            ## Get network element Flex Requirements
            network_element_df = pf_object.get_lines_constraints(network_element_df)

        elif specific_network_element_type == 'ElmTr2' or specific_network_element_type == 'ElmSite':

            network_element_df = pf_object.get_transformer_terminal_result(constrained_pf_id, loading_threshold_pct)

            ## Get network element Flex Requirements
            network_element_df = pf_object.get_trafos_constraints(network_element_df)

        oGen = pf_object.app.GetCalcRelevantObjects(flex_pf_id + '.ElmSym')[0]
        generator_name = oGen.loc_name
        generator_motor = bool(oGen.i_mot)
        generator_Pflex = oGen.pgini
        if generator_motor == False:
            generator_motor = 'generator'
            network_element_df = np.append(network_element_df, [generator_name, generator_motor])
            network_element_df = np.append(network_element_df, generator_Pflex)

        else:
            generator_motor = 'motor'
            network_element_df = np.append(network_element_df, [generator_name, generator_motor])
            network_element_df = np.append(network_element_df, generator_Pflex * (-1))

        network_element_df_all_steps.loc[iter_step] = network_element_df

    return network_element_df_all_steps


def PSA_residual_flex_only(status, pf_object, constrained_pf_id, constrained_pf_type, loading_threshold_pct):
    network_element_df = pd.DataFrame()
    if status == const.PSAok:
        ## Network element parameters
        specific_network_element = constrained_pf_id + '*.' + constrained_pf_type  # PowerFactory object of a certain type
        specific_network_element = pf_object.app.GetCalcRelevantObjects(specific_network_element)[0]

        if constrained_pf_type == 'ElmLne':

            ## Get network element results from loadflow
            network_element_df = pf_object.get_line_terminal_result(constrained_pf_id, loading_threshold_pct)

            ## Get network element Flex Requirements
            network_element_df = pf_object.get_lines_constraints(network_element_df)

        elif constrained_pf_type == 'ElmTr2':

            network_element_df = pf_object.get_transformer_terminal_result(constrained_pf_id, loading_threshold_pct)

            ## Get network element Flex Requirements
            network_element_df = pf_object.get_trafos_constraints(network_element_df)

    return network_element_df


def PSA_update_gen_run_ldf(pf_object, flex_pf_id, network_element_df_all_steps, iter_step):
    oGen = pf_object.app.GetCalcRelevantObjects(flex_pf_id + '.ElmSym')[0]

    if iter_step == 0:
        offered_power_kw = network_element_df_all_steps['P_req'][0] / 1e3
    else:
        offered_power_kw = network_element_df_all_steps['P_req'].sum() / 1e3

    oGen.pgini = offered_power_kw  # this offered power has to be converted to MW for PowerFactory

    print(f'Updated flex asset [{oGen.loc_name}] Pgini to [{offered_power_kw:.3f}]MW')

    print('=' * 65)
    ierr2 = pf_object.run_loadflow()
    if ierr2 == 0:
        status = const.PSAok
        print(
            f"step {iter_step}  \tLoad Flow SUCCESSFUL with Flex Asset [{flex_pf_id}] dispatched [{offered_power_kw:.3f}]MW")
    else:
        status = const.PSAPFError
        print(
            f"step {iter_step}  \tLoad flow ERROR with Flex Asset [{flex_pf_id}] dispatched [{offered_power_kw:.3f}]MW")
    print('=' * 65)

    return ierr2, status


def PSA_constr_resl(loading_diff_perc, stop_tolerance, status, pf_object,
                    constrained_pf_id, loading_threshold_pct, flex_pf_id,
                    network_element_df_all_steps):
    iter_step = 0

    while loading_diff_perc > stop_tolerance:
        print(f'iter_step:{iter_step} \t loading_diff_perc: {loading_diff_perc:.3f}')

        network_element_df_all_steps = PSA_residual_flex(status, pf_object,
                                                         constrained_pf_id, loading_threshold_pct,
                                                         flex_pf_id, iter_step, network_element_df_all_steps)

        ierr2, status = PSA_update_gen_run_ldf(pf_object, flex_pf_id, network_element_df_all_steps, iter_step)

        loading_diff_perc = network_element_df_all_steps['loading_diff_perc'].values[-1]
        iter_step = iter_step + 1

    final_loading = network_element_df_all_steps['Loading'].iloc[-1]
    final_loading_diff = network_element_df_all_steps['loading_diff_perc'].iloc[-1]
    print(
        f'FINISHED Constraint iteration loop, final loading: {final_loading:.3f}%, difference: {final_loading_diff:.3f}%')

    return network_element_df_all_steps

