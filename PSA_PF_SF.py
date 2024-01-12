import powerfactory_interface as pf_interface
import PSA_SND_Utilities as ut
import PSA_SND_Constants as const
import os
import pandas as pd
import PSA_SIA_Data_Input_Output as sio
import PSA_NeRDA_Data_Input_Output as nrd
import PSA_Calc_Flex_Reqts as calcflex
from datetime import datetime, timedelta
import math
import PF_Config as pfconf
import PSA_PF_Functions as pffun
import PSA_File_Validation as fv
import PSA_Events_Input_Output as eventsio
import PSA_Switch_Input_Output as switchio
import sys
from os.path import exists
import numpy as np
import UsefulPandas
import shutil

def PSA_PF_CalculateSF(df_network_element_all_steps):
    
    ## Get the direction of the flow on the network asset
    df_network_element_all_steps[['P_flow_sign',
                              'P_flow_dir']] = df_network_element_all_steps.apply(lambda x: get_direction_flow(x['P_LV'],
                                                                                        x['P_HV']), axis=1, result_type='expand')
    ## Get the sign of the power at the LV terminal of the network element
    df_network_element_all_steps['P_LV_sign'] = df_network_element_all_steps['P_LV'].apply(np.sign)

    ## Apply that sign to the MVA at the LV terminal of the network element
    df_network_element_all_steps['S_LV_w_sign'] = df_network_element_all_steps['S_LV'] * df_network_element_all_steps['P_LV_sign']

    ## Get the rolling difference of the apparent power at the network element LV terminal (absolute)
    df_network_element_all_steps['S_LV_diff_abs'] = running_diff_abs(
                np.array(df_network_element_all_steps['S_LV_w_sign']), 0, 1)

    ## Get the rolling difference of the active power at the network element LV terminal (absolute)
    df_network_element_all_steps['P_LV_diff_abs'] = running_diff_abs(np.array(df_network_element_all_steps['P_LV']),0, 1)

    ## Get the rolling difference of the reactive power at the network element LV terminal (absolute)
    df_network_element_all_steps['Q_LV_diff_abs'] = running_diff_abs(np.array(df_network_element_all_steps['Q_LV']),0, 1)

    ## Get the rolling difference of the active power of the flex asset  (absolute)                                                                            
    df_network_element_all_steps['Pflex_diff_abs'] = running_diff_abs(np.array(df_network_element_all_steps['Pflex']), 0, 1)
                
    ## Get the direction of the power flow difference in steps and the flow in P_LV_s0 at step 0
    df_network_element_all_steps['PF_direction_opp'] = df_network_element_all_steps.apply(
            lambda x: P_LV_diff_sign(x['P_LV_diff_abs'], df_network_element_all_steps['P_LV'].loc['initial']), axis=1)
    
    ## Calculate the sensitivity factor of d_MVA_netasset/d_MW_flex
    df_network_element_all_steps['SF_abs'] = df_network_element_all_steps.apply(
            lambda x: delta_MVA_delta_MW(x['S_LV_diff_abs'],x['Pflex_diff_abs']),axis=1)

    ## Get the sensitivity factor direction, given by Delta_Pflow (flow in the network element) - absolute
    df_network_element_all_steps['SF_abs_sign'] = df_network_element_all_steps['P_LV_diff_abs'].apply(pd_npsign_alt)

    ## Validate the sensitivity factor by applying the sensitivity factor to the Pflex and add/subtract the initial S_LV with sign (S_LV_w_sign_s0)
    df_network_element_all_steps['SF_validation'] = df_network_element_all_steps.apply(lambda x: SF_validation(
                                                                        x['SF_abs'],
                                                                        x['SF_abs_sign'],
                                                                        x['PF_direction_opp'],
                                                                        df_network_element_all_steps['P_LV'].loc['initial'],
                                                                        x['Pflex'],
                                                                        df_network_element_all_steps['S_LV_w_sign'].loc['initial']), axis=1)

    df_network_element_all_steps['SF_validation_check'] = df_network_element_all_steps['SF_validation'] == df_network_element_all_steps['S_LV']
    df_network_element_all_steps['PF_LV'] = abs(df_network_element_all_steps['P_LV'])/df_network_element_all_steps['S_LV']
    df_network_element_all_steps['PF_LV_diff_abs'] = running_diff_abs(np.array(df_network_element_all_steps['PF_LV']), 0, 1)

    sensitivity_factor= 0 if math.isnan(df_network_element_all_steps['SF_abs']['final']) \
        else df_network_element_all_steps['SF_abs']['final']
    sensitivity_factor_direction = 0 if math.isnan(df_network_element_all_steps['SF_abs_sign']['final']) \
        else df_network_element_all_steps['SF_abs_sign']['final']

    power_factor_LV=df_network_element_all_steps['PF_LV']['final']

    print(f'Sensitivity Factor: {sensitivity_factor:.3f}')
    print(f'Sensitivity Factor direction: {sensitivity_factor_direction}')
    print(f'Power Factor at LV side: {power_factor_LV:.3f}')

    return df_network_element_all_steps, sensitivity_factor, sensitivity_factor_direction


def get_direction_flow(P_LV, P_HV):
    '''This function gives you the direction of the flow in a network element'''
    if  P_HV > 0:
        sign = np.sign(P_HV)
        direction = 'Import'
    else:
        sign = np.sign(P_HV)
        direction = 'Export'
    return sign, direction


def delta_MVA_delta_MW(mva, mw):
    if abs(mva) > 0 and abs(mw) > 0:
        return abs(mva) / abs(mw)
    else:
        return np.nan

def running_diff_abs(arr, ref_start, N):
    '''Return the rolling difference with respect to the first value i.e. absolute'''
    return np.array([arr[i - N] - arr[ref_start] for i in range(N, len(arr) + 1)])

def SF_validation(SF, SF_sign, PF_direction_opp, P_LV_s0, Pflex, S_LV):
    ## If the direction of the powerflow at the LV terminal of the network element at the current step (np.sign(Ps0) == np.sign(Psx-Ps0) is the same
    ## as the direction of the flow of the initial step np.sign(Ps0)
    if (PF_direction_opp == True and np.sign(P_LV_s0) == -1):
        # if (PF_direction_opp==False):
        SF_val = abs(SF * SF_sign * Pflex + S_LV)
    else:
        SF_val = abs(SF * SF_sign * Pflex - S_LV)

    return SF_val

def P_LV_diff_sign(P_LV_diff_abs, P_LV):
    '''This function gives you the direction of the flow between
    the difference in steps and the flow in P_LV_s0 at step 0 (this is always the reference)'''
    if P_LV_diff_abs == 0:
        pf_opposite = np.nan
    elif np.sign(P_LV_diff_abs) == np.sign(P_LV):
        pf_opposite = False
    else:
        pf_opposite = True

    return pf_opposite

def pd_npsign_alt(num):
    if abs(num) > 0:
        return np.sign(num)
    else:
        return np.nan

def PSA_SF_Responses_table(reqid, resp_id, bsp, primary, feeder, secondary, terminal, constrained_pf_id,
                           constrained_pf_type, asset_loading_initial, asset_loading_final, busbar_from, busbar_to, scenario, start_time, duration,required_power_kw,
                             flex_pf_id, offered_power_kw,calculate_historical,sf_calculated_datetime,
                           response_datetime,entered_datetime,
                           sensitivity_factor,sensitivity_factor_direction):
    data = [{"req_id": reqid, "resp_id": resp_id, "bsp": bsp, "primary": primary, "feeder": feeder,
             "secondary": secondary, "terminal": terminal, "constrained_pf_id": constrained_pf_id,
             "constrained_pf_type": constrained_pf_type, "asset_loading_initial": asset_loading_initial,
             "asset_loading_final": asset_loading_final, "busbar_from":busbar_from, "busbar_to": busbar_to,
             "scenario": scenario,"start_time": start_time, "duration": duration,
             "required_power_kw": required_power_kw,"flex_pf_id": flex_pf_id, "offered_power_kw": offered_power_kw,
             "calculate_historical": calculate_historical,"sf_calculated_datetime": sf_calculated_datetime,
             "response_datetime": response_datetime, "entered_datetime": entered_datetime,
             "sensitivity_factor":sensitivity_factor, "sensitivity_factor_dir":sensitivity_factor_direction}]
    dfPSASFResponses = pd.DataFrame(data)
    return dfPSASFResponses

def PSA_PF_readSIA(SIAfolder):
    dictSIAFeeder = dict()
    dictSIAGen = dict()
    dictSIAGroup = dict()
    SIAfiles = os.listdir(SIAfolder)
    for SIAfile in SIAfiles:
        if SIAfile.startswith("SIA_Feeder"):
            SIAFeederXls = pd.ExcelFile(SIAfolder + "\\" + SIAfile)
            lSIAParams = SIAFeederXls.sheet_names
            Feedername = SIAfile.replace("SIA_Feeder_", "")
            dictSIAFeeder[Feedername] = dict()
            for SIAParam in lSIAParams:
                dictSIAFeeder[Feedername][SIAParam] = SIAFeederXls.parse(SIAParam, index_col=0)
        elif SIAfile.startswith("SIA_Gen"):
            SIAGenXls = pd.ExcelFile(SIAfolder + "\\" + SIAfile)
            lSIAParams = SIAGenXls.sheet_names
            Genname = SIAfile.replace("SIA_Gen_", "")
            dictSIAGen[Genname] = dict()
            for SIAParam in lSIAParams:
                dictSIAGen[Genname][SIAParam] = SIAGenXls.parse(SIAParam, index_col=0)
        elif SIAfile.startswith("SIA_Group"):
            SIAGroupXls = pd.ExcelFile(SIAfolder + "\\" + SIAfile)
            lSIAParams = SIAGroupXls.sheet_names
            Groupname = SIAfile.replace("SIA_Group_", "")
            dictSIAGroup[Groupname] = dict()
            for SIAParam in lSIAParams:
                dictSIAGroup[Groupname][SIAParam] = SIAGroupXls.parse(SIAParam, index_col=0)
    return dictSIAFeeder, dictSIAGen, dictSIAGroup


def downscale_SIA_to_loads(all_loads_df, dictSIAFeeder, i_day, i_half_hour, dfDownScalingFactor_grid, power_factor,
                           lagging, verbose):
    ## Load SIA load data for all loads ##
    s_load_day_hh_dict = {}
    for strNameFeeder, dictFeeder in dictSIAFeeder.items():
        dfSIAFeeder = dictFeeder['UND_DEM_MVA']
        strNameFeeder = strNameFeeder.replace(".xlsx", "")
        s_load_day_hh = dfSIAFeeder.iloc[i_half_hour, i_day]
        s_load_day_hh_dict[strNameFeeder] = s_load_day_hh
    if verbose == True:
        print(s_load_day_hh_dict)

    ## Map the s_load_day_hh_dict into the existing downscaling dataframe
    dfDownScalingFactor_grid['S_load_new'] = dfDownScalingFactor_grid["Grid"].map(s_load_day_hh_dict)
    ## Downscale S_load_new
    dfDownScalingFactor_grid["S_indv_setpoint"] = dfDownScalingFactor_grid["S_load_new"] * (
            dfDownScalingFactor_grid["S_rated_indv_perc"] / 100)
    ## Calculate P_indv_setpoint and Q_indv_setpoint
    dfDownScalingFactor_grid['P_indv_setpoint'] = dfDownScalingFactor_grid['S_indv_setpoint'].apply(
        lambda x: x * power_factor)

    if lagging:
        dfDownScalingFactor_grid['Q_indv_setpoint'] = dfDownScalingFactor_grid['S_indv_setpoint'].apply(
            lambda x: x * np.sin(np.arccos(power_factor)))
    else:
        dfDownScalingFactor_grid['Q_indv_setpoint'] = -dfDownScalingFactor_grid[
            'S_indv_setpoint'].apply(lambda x: x * np.sin(np.arccos(power_factor)))

    for index, row in dfDownScalingFactor_grid.iterrows():

        load_name = row['Name_load']
        load_name_grid = row['Name_grid']
        load_obj = row['PF_object_load']
        load_p_nom = row['P_nom']
        load_q_nom = row['Q_nom']
        load_s_new = row['S_indv_setpoint']
        load_p_new = row['P_indv_setpoint']
        load_q_new = row['Q_indv_setpoint']

        if verbose == True:
            print(
                f"{index + 1}/{dfDownScalingFactor_grid.shape[0]} - [{load_name}], [{load_name_grid}], [{load_p_nom:.3f}], [{load_q_nom:.3f}], | [{load_p_new:.3f}], [{load_q_new:.3f}]")

        ### Look for the PowerFactory object first (look in the dataframe rather than PF)
        load_obj = \
            UsefulPandas.filter_df_single(all_loads_df, 'Name_grid', load_name_grid, regex=None)[
                'PF_object'].values[0]

        ### Then assign the P/Q setpoints to the PF object
        load_obj.plini = row['P_indv_setpoint']
        load_obj.qlini = row['Q_indv_setpoint']

    if verbose == True:
        print('Done changing P/Q setpoints of load objects according to forecast and downscaling')

    return all_loads_df

def PSA_PF_runningmode(dict_params):
    if dict_params[const.strMaintFile] == "NULL" and dict_params[const.strContFile] == "NULL":
        PSA_running_mode = "B"
    elif dict_params[const.strMaintFile] == "NULL" and dict_params[const.strContFile] != "NULL":
        PSA_running_mode = "C"
    elif dict_params[const.strMaintFile] != "NULL" and dict_params[const.strContFile] == "NULL":
        PSA_running_mode = "M"
    else:
        PSA_running_mode = "MC"
    return PSA_running_mode

def PSA_PF_find_datetime_idx(psa_start_datetime,constraint_start_datetime, verbose):
    
    psa_start_datetime=pd.to_datetime(psa_start_datetime,infer_datetime_format=True)
    constraint_start_datetime=pd.to_datetime(constraint_start_datetime,infer_datetime_format=True)
    status = const.PSAok

    ## Shift to midnight
    psa_start_datetime_midnight = psa_start_datetime.replace(hour=0, minute=0, second=0, microsecond=0)
    
    ## Calculate the time delta between the two datetime objects
    ## future minus past
    delta=constraint_start_datetime-psa_start_datetime_midnight
    
    # Calculate the time delta in days, hours and half hours
    delta_tot_seconds=delta.total_seconds()
    days, seconds = divmod(delta_tot_seconds, 86400)
    half_hours, seconds_res = divmod(seconds, 1800)
    minutes, seconds_res = divmod(seconds, 60)
    
    ## Get the number of half hours including decimals
    half_hours_float=minutes/30
    
    i_day=int(days)

    if verbose == True:
        print(f'psa_start_datetime: {psa_start_datetime}')
        print(f'psa_start_datetime_midnight: {psa_start_datetime_midnight}')
        print(f'constraint_start_datetime: {constraint_start_datetime}')
    
    ## Check whether the delta is positive (i.e. we are subtracting future from past)
    if float(delta_tot_seconds) < 0:
        status = const.PSAfileReadError
        if verbose == True:
            print(f'The constraint_start_datetime should be in the future!! \n{constraint_start_datetime}')
    else:
        if verbose == True:
            print(f'The constraint is {i_day} days, {half_hours} half hours ahead of the PSA run from midnight')
    
    residual_half_hour, i_half_hour = math.modf(half_hours_float)
    
    if residual_half_hour>0.5:
        i_half_hour=i_half_hour+1
        
    i_half_hour=int(i_half_hour) ### a -1 might be needed if the index starting at zero, however NOT needed if midnight shift
    if verbose == True:
        print(f'day:[{i_day}], half-hours flt:[{half_hours_float:.2f}], minutes:[{minutes}], i_half_hour:[{i_half_hour}]')
    return status, i_half_hour, i_day

def getFilePSArunID(fname):
    # if bSF:
    #     vn = os.path.basename(fname).split('-')[-1].replace('.xlsx', '')
    # else:
    #     vList = os.path.basename(fname).split('-')
    #     if vList[-1].find("V") == -1:
    #         vn = '-'.join(os.path.basename(fname).split('-')[-2:]).replace('.xlsx', '')
    #     else:
    #         vn = vList[-1].replace('.xlsx', '')
    filename = os.path.basename(fname)
    filename_without_ext = os.path.splitext(filename)[0]  # Remove the file extension
    filename_parts = filename_without_ext.split("-")  # Split the filename into its parts
    version_string = [part for part in filename_parts if part.startswith("V")][
        -1]  # Get the last part that starts with "V"
    version_index = filename_parts.index(version_string)
    if version_index == len(filename_parts) - 2:
        vn = '-'.join([version_string, filename_parts[-1]])
    else:
        vn = version_string
    print("\n################################################################")
    print(os.path.basename(fname))
    print(f'The SND responses/contracts file version number is: {vn}')
    PSArunID = os.path.basename(fname)[:20]
    PSAfolder = fname.replace(os.path.basename(fname), '')
    print(f"PSArunID: [{PSArunID}]")
    print(f"results_folder: [{PSAfolder}]")
    print("################################################################\n")
    return vn, PSArunID, PSAfolder

def getPSArunIDdate(PSArunID):
    strDate = PSArunID[4:-6]
    return strDate

def getPSArunIDtime(PSArunID):
    strTime = PSArunID[15:]
    return strTime

def PSA_PF_workflowSF(bOutput2SND, bAuto, SND_file, verbose=False):
    print("\n################################################################")
    print('Calculating sensitivity factors')
    print("################################################################\n")
    print(f'bOutput2SND; [{bOutput2SND}]')
    print(f'bAuto; [{bAuto}]')
    print(f'SND_file; [{SND_file}]')
    print('-'*65)

    status = const.PSAok
    msg = ""
    dict_config = ut.PSA_SND_read_config(const.strConfigFile)
    # get PSArunID from filename
    vn, PSArunID, PSAfolder = getFilePSArunID(SND_file)
    dict_params = ut.readPSAparams(PSAfolder + PSArunID + const.strRunTimeParams)
    if dict_params[const.strMaint] == "True" and not dict_params[const.strCont] == "False":
        scenario = "MAINT"
    elif dict_params[const.strMaint] == "False" and not dict_params[const.strCont] == "True":
        scenario = "CONT"
    elif dict_params[const.strMaint] == "True" and not dict_params[const.strCont] == "True":
        scenario = "MAINT_CONT"
    else:
        scenario = "BASE"

    dict_config = ut.PSA_SND_read_config(const.strConfigFile)
    
    # if in Auto mode being run by S&D process and assume multiprocess
    if bAuto:
        multiProcess = True
    else:
        multiProcess = False

    pf_object, status, msg = pf_interface.run(dict_config[const.strPFUser], multiProcess)
    if status != const.PSAok:
        msg = "PSA_PF_SF: " + msg
        return status, msg
    pf_interface.activate_project(pf_object, dict_params[const.strBSPModel],
                                    os.path.join(str(PSAfolder),str(const.strINfolder).replace("\\","")))
       
    if (not bAuto) and pfconf.bCreateRef:       
        outputFile = os.path.join(dict_config[const.strWorkingFolder], const.strRefAssetFile)
        pffun.PF_getMasterRefList(pf_object, outputFile)
        print("Reference list of PowerFactory Assets has been updated!")
    
    # validate asset names
    lRefLine, lRefTrafo, lRefSyncMachine, lRefSwitch, lRefCoupler = fv.readRefAssetNames(os.path.join(dict_config[const.strWorkingFolder], const.strRefAssetFile))
    status,msg = fv.validateSNDFileNames(lRefSyncMachine, SND_file)
    
    if status != const.PSAok:
        return status, msg
    
    # if bAuto:
    #     # PSArunID = SND_file[0:23]
    #     results_folder = dict_config[const.strWorkingFolder] + "\\" + PSArunID + "\\"
    #     print("PSArunID = " + PSArunID)
    #     print("results_folder = " + results_folder)
    # else:
    #     PSArunID, results_folder = getFilePSArunID(SND_file)
        
    # PSAfolder = results_folder + "\\"
    start_date = getPSArunIDdate(PSArunID)
    start_time = getPSArunIDtime(PSArunID)
    
    if SND_file.find(const.strSNDResponses) >=0:
        bResponse=True
    elif SND_file.find(const.strSNDContracts) >=0:
        bResponse = False
    else:
        status = const.PSAfileTypeError
        msg = "Unknown file type: " + SND_file
        return status, msg

    ### Load PSArunID folder ###
    if exists(PSAfolder + PSArunID + const.strRunTimeParams):
        ### Load runtime parameters ###   
        dict_params = ut.readPSAparams(PSAfolder + PSArunID + const.strRunTimeParams)
        print("Loaded PSArunID parameters")
    else:
        status = const.PSAfileExistError
        msg = "Run time Params file doesn't exist: " + PSAfolder + PSArunID + const.strRunTimeParams
        return status, msg

    ### Load input parameters ###
    bDebug = pfconf.bDebug
    number_of_days = int(dict_params[const.strDays])
    number_of_half_hours = int(dict_params[const.strHalfHours])
    power_factor = float(dict_params[const.strPowFctr])
    lagging = (str(dict_params[const.strPowFctrLag]) == "True")
    lDfPSASFResponses = []
    file_detector_folder = dict_config[const.strFileDetectorFolder]
    print("BSP Model = " + dict_params[const.strBSPModel])

    # CREATE TEMP SF FOLDER
    SF_temp_file_path = PSAfolder + "\\" + const.strTempSFfiles
    if not os.path.isdir(SF_temp_file_path):
        os.mkdir(SF_temp_file_path)
    SF_temp_file_path_vn = SF_temp_file_path + "\\" + vn
    if not os.path.isdir(SF_temp_file_path_vn):
        os.mkdir(SF_temp_file_path_vn)

    # LOAD DATA FROM THE PSArunID FOLDER NOT THE SND FILE DETECTOR FOLDER
    dfSNDoutput = pd.read_excel(SND_file)
    dfSNDoutput_copy = dfSNDoutput.copy()
    dfSNDoutput = ut.aggregateSNDresponses(dfSNDoutput)
    dfSNDoutput.to_excel(SF_temp_file_path_vn+"\\"+PSArunID+const.strSNDResponses+"-DEDUP"+"-"+vn+".xlsx", index=False)
    fname_AssetData = dict_params[const.strAssetData]
    calculate_historical = dfSNDoutput['calculate_historical'][0]
    input_folder = PSAfolder + const.strINfolder
    pf_proj_dir = input_folder

    if dict_params[const.strbEvent] == "True":
        bEvent = True
    else:
        bEvent = False

    if dict_params[const.strSwitch] == "True":
        bSwitch = True
    else:
        bSwitch = False

    if bEvent:
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

    dfSIAData = pd.read_excel(os.path.join(input_folder, fname_AssetData), sheet_name=const.strSIASheet)
    dfNeRDAData = pd.read_excel(os.path.join(input_folder, fname_AssetData), sheet_name=const.strNeRDASheet)

    start_date_time = start_date + " " + start_time.split("-")[0] + ":" + start_time.split("-")[1] + ":00"

    ### Load SIA data based on calculate_historical ###
    SIAfolder = PSAfolder+"\\" + const.strSIAfolder
    SIAoutputfolder = PSAfolder + "\\" + const.strNewSIAfolder
    if not os.path.isdir(SIAoutputfolder):
        os.mkdir(SIAoutputfolder)

    if calculate_historical:
        dictSIAFeeder, dictSIAGen, dictSIAGroup = PSA_PF_readSIA(SIAfolder)
        SIA_start_date_time = start_date_time
    else:
        token, status, msg = sio.SIA_get_token(dict_config[const.strSIAUsr], dict_config[const.strSIAPwd],
                                               int(dict_config[const.strSIATimeOut]))
        if status != const.PSAok:
            return status, msg
        else:
            SIA_start_time, SIA_start_date_time, status, msg, dictSIAFeeder, dictSIAGen, dictSIAGroup = sio.loadSIAData(
                strbUTC=dict_config[const.bUTC], assets=dfSIAData, timeout=int(dict_config[const.strSIATimeOut]),
                OutputFilePath=SIAoutputfolder, token=token,bDebug=bDebug,PSArunID=PSArunID,
                number_of_days=number_of_days, number_of_half_hours=number_of_half_hours)
            if status != const.PSAok:
                return status, msg

    ### Load NeRDA data based on calculate_historical ###
    NeRDAfolder = PSAfolder+"\\"+"2 - NeRDA_DATA"

    if calculate_historical:
        copyNeRDAfile1 = dict_config[const.strWorkingFolder] + "\\" + const.strDataResults + "\\AUT_2023-MM-DD-HH-MM-NeRDA_DATA.xlsx"        
        copyNeRDAfile2 = NeRDAfolder + "\\AUT_2023-MM-DD-HH-MM-NeRDA_DATA.xlsx"
        shutil.copyfile(copyNeRDAfile1, copyNeRDAfile2)
        NeRDAfiles = os.listdir(NeRDAfolder)
        for NeRDAfile in NeRDAfiles:
            if NeRDAfile.endswith("NeRDA_DATA.xlsx"):
                dfNeRDA = pd.read_excel(NeRDAfolder + "\\" + NeRDAfile)

    else:
        if pfconf.bUseNeRDA:
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
        calcflex.PSA_update_switch(pf_object, dfElmCoup_ALL, dfStaSwitch_ALL, dfNeRDACopy)

    ### Read downscaling factors ###
    ### OLD CODE ###
    #if dict_params[const.strAuto] == "True":
    #    strDownscalingFactorFolder = PSAfolder + const.strBfolder + "\\DOWNSCALING_FACTOR\\"
    #elif dict_params[const.strMaint] == "True" and dict_params[const.strCont] == "False":
    #    strDownscalingFactorFolder = PSAfolder + const.strMfolder + "\\DOWNSCALING_FACTOR\\"
    #elif dict_params[const.strMaint] == "False" and dict_params[const.strCont] == "True":
    #    strDownscalingFactorFolder = PSAfolder + const.strCfolder + "\\DOWNSCALING_FACTOR\\"
    #elif dict_params[const.strMaint] == "True" and dict_params[const.strCont] == "True":
    #    strDownscalingFactorFolder = PSAfolder + const.strMCfolder + "\\DOWNSCALING_FACTOR\\"

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
                dfDownScalingFactor_grid = pd.read_excel(os.path.join(strDownscalingFactorFolder,
                                                                    'ALL_PSA_DOWNSCALING_FACTORS_PQNOM.xlsx'))
                file_found = True

    if file_found == False:
        status = const.PSAfileReadError
        msg = f"Donwscaling Factor file not found in {strDownscalingFactorFolder}"
        return status, msg
    else:
        dfDownScalingFactor_grid['Name_grid']=dfDownScalingFactor_grid['Name']+'_'+dfDownScalingFactor_grid['Grid']

    ### Get the list of lines and trafos in the network ###
    all_loads_df = pf_object.get_loads_names_grid_obj()
    ### List of scenarios ###
    groupedbyscenario = dfSNDoutput.groupby(['scenario'])
    lScenarios = list(set(dfSNDoutput['scenario']))

    ### Loop through scenarios ###
    for idx_sc, scenario in enumerate(lScenarios):
        print(f"================================\n {idx_sc+1}/{len(lScenarios)} Scenario: [{scenario}]\n================================")
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
        if scenario == "MAINT" or scenario == "CONT" or scenario == "MAINT_CONT":
            dictDfTrafos_2w, dictDfLine, dictDfSynm,dictDfLoad = ut.readOutageFile(scenario, PSAfolder)
        dfScenario = groupedbyscenario.get_group(scenario)
        ### Loop through start_time ###
        groupedbystarttime = dfScenario.groupby(['start_time'])
        lStarttime = list(set(dfScenario['start_time']))
        if verbose == True:
            print("================================================")
            print(dfScenario.shape)
            print("lStarttime are:", lStarttime)
            print("================================================")
        for idx_stt, starttime in enumerate(lStarttime):
            time_diff = (datetime.strptime(starttime, '%Y-%m-%d %H:%M:%S') -
                         datetime.strptime(SIA_start_date_time, '%Y-%m-%d %H:%M:%S')).days
            if time_diff < 0:
                ispast = True
            else:
                ispast = False
                print(f"Constraint happened at {starttime} is in the past; it is skipped.")
            if ispast == False:
                print(f"================================\n {idx_stt + 1}/{len(lStarttime)} Constrained time: [{starttime}]")
                dfStarttime = groupedbystarttime.get_group(starttime)
                ### Loop through req_ids ###
                groupedbyreqid = dfSNDoutput.groupby(['req_id'])
                setReqids = list(set(dfStarttime['req_id']))
                ### Read the start_time of the req_ids with the same start time ###
                if verbose == True:
                    print("================================================")
                    print(dfStarttime.shape)
                    print("setReqids are:", setReqids)
                    print("================================================")
                status, i_half_hour, i_day = PSA_PF_find_datetime_idx(start_date_time, starttime, verbose=False)
                print("###############################")
                print("===============================")
                print(f"i_day is {i_day}")
                print(f"SIA_start_date_time is {datetime.strptime(SIA_start_date_time, '%Y-%m-%d %H:%M:%S')}")
                print(f"start_date_time is {datetime.strptime(start_date_time, '%Y-%m-%d %H:%M:%S')}")
                print("===============================")
                print("###############################")

                i_day_shifted = i_day - (datetime.strptime(SIA_start_date_time, '%Y-%m-%d %H:%M:%S').replace(hour=0, minute=0, second=0, microsecond=0) -
                                         datetime.strptime(start_date_time, '%Y-%m-%d %H:%M:%S').replace(hour=0, minute=0, second=0, microsecond=0)).days
                if scenario == "MAINT" or scenario == "CONT" or scenario == "MAINT_CONT":
                    pffun.PF_update_outages(pf_object, dictDfLine, dictDfTrafos_2w, dictDfSynm, dictDfLoad, i_half_hour, i_day, verbose=True)
                    print(f'\t --Done updating outages for {scenario} {starttime}')
                if verbose == True:
                    print(f'day:[{i_day}], i_half_hour:[{i_half_hour}]')
                if status == const.PSAok:
                    ## Load Events data ##
                    if bEvent:
                        eventsio.updateServices(pf_object, dictDfServices, i_half_hour, i_day, verbose=True)
                    if bSwitch:
                        switchio.updateSwitchStatus(pf_object, dictSwitchStatus, i_half_hour, i_day, verbose=True)
                        print("-" * 20)
                        print(f"Switch status is updated for day {i_day} and half hour {i_half_hour}.")
                        print("-" * 20)
                    ## Load SIA gen data for non- _F assets that have SIA data ##
                    for Genname, dictSIAGenParams in dictSIAGen.items():
                        Genname = Genname.replace(".xlsx", "")
                        oGen = pf_object.app.GetCalcRelevantObjects(Genname + '.ElmSym')[0]
                        oGen.pgini = dictSIAGenParams['GEN_MW'].iloc[i_half_hour][i_day_shifted]
                    if verbose == True:
                        print('Done updating Generator setpoints according to SIA')
                    all_loads_df = downscale_SIA_to_loads(all_loads_df, dictSIAFeeder, i_day_shifted, i_half_hour,
                                                          dfDownScalingFactor_grid, power_factor, lagging,verbose=False)
                    print('Done changing P/Q setpoints of load objects according to forecast and downscaling')

                    ### Iterate through req_ids ###
                    for idx_rq,reqid in enumerate(setReqids):
                        dfReqid = groupedbyreqid.get_group(reqid)
                        constrained_pf_id = dfReqid.iloc[0]['constrained_pf_id']
                        constrained_pf_type = dfReqid.iloc[0]['constrained_pf_type']
                        print(f"\n{idx_rq + 1}/{len(setReqids)} - Requirement ID {reqid} - day:[{i_day}] hh:[{i_half_hour}] - constrained_pf_id: {constrained_pf_id}")

                        ### Prepare load flow ###
                        pf_object.prepare_loadflow(ldf_mode=dict_config[const.strPFLFMode], algorithm=dict_config[const.strPFAlgrthm],
                                                   trafo_tap=dict_config[const.strPFTrafoTap],shunt_tap=dict_config[const.strPFShuntTap],
                                                   zip_load=dict_config[const.strPFZipLoad], q_limits=dict_config[const.strPFQLim],
                                                   phase_shift=dict_config[const.strPFPhaseShift],
                                                   trafo_tap_limit=dict_config[const.strPFTrafoTapLim],
                                                   max_iter=int(dict_config[const.strPFMaxIter]))

                        ### Iterate through resp_ids as each row of req_ID dataframe which is individual response ###
                        i = 0
                        for idx_rp, row in dfReqid.iterrows():
                            ## obtain %loading for overloaded asset ##
                            i += 1
                            print('='*65)
                            if bResponse:
                                resp_id = row['resp_id']
                            else:
                                resp_id = row['con_id']
                            bsp = row['bsp']
                            primary = row['primary']
                            feeder = row['feeder']
                            secondary = row['secondary']
                            terminal = row['terminal']
                            required_power_kw = row['required_power_kw']
                            busbar_from = row['busbar_from']
                            busbar_to = row['busbar_to']
                            start_time = row['start_time']
                            duration = row['duration']
                            response_datetime = row['response_datetime']
                            entered_datetime = row['entered_datetime']
                            flex_pf_id = row['flex_pf_id']
                            offered_power_mw = row['offered_power_kw']/1e3 #this offered power has to be converted to MW for PowerFactory
                            if bResponse:
                                print(
                                    f"{i}/{len(dfReqid['resp_id'])} - Response ID [{resp_id}] - flex asset: {flex_pf_id}")
                            else:
                                print(
                                    f"{i}/{len(dfReqid['con_id'])} - Contract ID [{resp_id}] - flex asset: {flex_pf_id}")
                            ## Run Load Flow For Each Resp_ID##
                            print('=' * 65)
                            ierr1 = pf_object.run_loadflow()
                            if ierr1 == 0:
                                status = const.PSAok
                                print(
                                    f"Load Flow SUCCESSFUL for Initial Conditions (before flex) - Day: {i_day}, Half Hour:{i_half_hour} - {starttime}")
                            else:
                                status = const.PSAPFError
                                print(
                                    f"Load flow ERROR for Initial Conditions (before flex) - Day: {i_day},  Half Hour:{i_half_hour} - {starttime}")
                            print('=' * 65)

                            ## Obtain the power flow results for initial network condition before flex asset dispatch ##
                            network_element_df_all_steps = pd.DataFrame(columns=['PF_Object_network_element',
                                                                                 'Name_network_element',
                                                                                 'Type', 'Grid',
                                                                                 'From', 'To',
                                                                                 'From_Status', 'To_Status',
                                                                                 'P_HV', 'Q_HV', 'S_HV', 'I_HV', 'V_HV',
                                                                                 'P_LV', 'Q_LV', 'S_LV', 'I_LV', 'V_LV',
                                                                                 'I_nom_act_kA_lv', 'I_nom_act_kA_hv',
                                                                                 'I_max_act',
                                                                                 'loading_threshold_pct',
                                                                                 'Loading_pct', 'loading_diff_perc',
                                                                                 'Ploss', 'Qloss', 'Overload',
                                                                                 'OutService', 'CIM_ID', 'Service',
                                                                                 'Name_flex_asset', 'Generator_motor',
                                                                                 'Pflex'
                                                                                 ])
                            if status == const.PSAok:
                                specific_network_element = constrained_pf_id + '*.' + constrained_pf_type  # PowerFactory object of a certain type
                                specific_network_element = pf_object.app.GetCalcRelevantObjects(specific_network_element)[0]
                                specific_network_element_name = specific_network_element.loc_name

                                if constrained_pf_type == 'ElmLne':
                                    network_element_df = pf_object.get_line_terminal_result(constrained_pf_id, 100)

                                elif constrained_pf_type == 'ElmTr2':
                                    network_element_df = pf_object.get_transformer_terminal_result(constrained_pf_id, 100)

                                else:
                                    network_element_df = pd.DataFrame()
                                oGen_flex = pf_object.app.GetCalcRelevantObjects(flex_pf_id + '.ElmSym')[0]
                                generator_name = oGen_flex.loc_name
                                generator_motor = bool(oGen_flex.i_mot)
                                generator_Pflex = oGen_flex.pgini
                                if generator_motor == False:
                                    generator_motor = 'generator'
                                    network_element_df = np.append(network_element_df, [generator_name, generator_motor])
                                    network_element_df = np.append(network_element_df, generator_Pflex)

                                else:
                                    generator_motor = 'motor'
                                    network_element_df = np.append(network_element_df, [generator_name, generator_motor])
                                    network_element_df = np.append(network_element_df, generator_Pflex * (-1))
                                network_element_df_all_steps.loc['initial'] = network_element_df

                            else:
                                asset_loading = 0

                            ## Change Flex Asset Dispatch For Each Resp_ID ##
                            oGen_flex = pf_object.app.GetCalcRelevantObjects(flex_pf_id+'.ElmSym')[0]
                            oGen_flex.pgini = offered_power_mw # Update the dispatch power of the flex asset
                            print(f'Updated flex asset [{oGen_flex.loc_name}] Pgini to [{offered_power_mw}]MW')

                            ## Run load flow ##
                            print('='*65)
                            ierr2 = pf_object.run_loadflow()
                            if ierr2 == 0:
                                status = const.PSAok
                                print(f"Load Flow SUCCESSFUL with Flex Asset [{flex_pf_id}] dispatched (after flex) - Day: {i_day}, Half Hour:{i_half_hour} - {starttime}")
                            else:
                                status = const.PSAPFError
                                print(f"Load flow ERROR with Flex Asset [{flex_pf_id}] dispatched (after flex) - Day: {i_day}, Half Hour:{i_half_hour} - {starttime}")
                            print('='*65)

                            ## Obtain loading for load flow with redispatch of Flex Asset ##
                            if status == const.PSAok:
                                specific_network_element = constrained_pf_id + '*.' + constrained_pf_type
                                specific_network_element = pf_object.app.GetCalcRelevantObjects(specific_network_element)[0]
                                specific_network_element_name = specific_network_element.loc_name
                                network_element_df = pd.DataFrame()
                                if constrained_pf_type == 'ElmLne':
                                    network_element_df = pf_object.get_line_terminal_result(constrained_pf_id, 100)

                                elif constrained_pf_type=='ElmTr2':
                                    network_element_df = pf_object.get_transformer_terminal_result(constrained_pf_id, 100)

                                generator_name = oGen_flex.loc_name
                                generator_motor = bool(oGen_flex.i_mot)
                                generator_Pflex=oGen_flex.pgini
                                if generator_motor == False:
                                    generator_motor = 'generator'
                                    network_element_df = np.append(network_element_df, [generator_name, generator_motor])
                                    network_element_df = np.append(network_element_df, generator_Pflex)

                                else:
                                    generator_motor = 'motor'
                                    network_element_df = np.append(network_element_df, [generator_name, generator_motor])
                                    network_element_df = np.append(network_element_df, generator_Pflex * (-1))
                                network_element_df_all_steps.loc['final']=network_element_df
                                network_element_df_all_steps['PF_Object_network_element']=\
                                    network_element_df_all_steps['PF_Object_network_element'].apply(str)

                                csvFile = SF_temp_file_path_vn + "\\" + f'SF_int_{flex_pf_id}_{specific_network_element_name}_{reqid}_{resp_id}.csv'
                                if len(csvFile) >= 255:
                                    csvFile = SF_temp_file_path_vn + "\\" + f'SF_int_{flex_pf_id[:5]}_{specific_network_element_name[-7:]}_{reqid}_{resp_id}.csv'
                                network_element_df_all_steps.to_csv(csvFile, index=True)

                                #network_element_df_all_steps.to_csv(os.path.join(SF_temp_file_path,
                                #   f'SF_int_{flex_pf_id}_{specific_network_element_name}_{reqid}_{resp_id}.csv'),index=True)
                                network_element_df_all_steps_1 = network_element_df_all_steps.copy()
                                network_element_df_all_steps_1, sensitivity_factor, sensitivity_factor_direction = PSA_PF_CalculateSF(network_element_df_all_steps_1)
                                asset_loading_initial = float(network_element_df_all_steps_1['Loading_pct']['initial'])
                                asset_loading_final = float(network_element_df_all_steps_1['Loading_pct']['final'])
                                print(f"asset_loading_final is {asset_loading_final}")
                                network_element_df_all_steps_2=network_element_df_all_steps_1[['Name_network_element', 'Grid',
                                                                                        'S_LV','P_LV','Q_LV','PF_LV',
                                                                                        'Loading_pct','Name_flex_asset',
                                                                                        'Generator_motor', 'Pflex',
                                                                                        'P_flow_sign', 'P_flow_dir', 'P_LV_sign',
                                                                                        'PF_direction_opp', 'PF_LV_diff_abs',
                                                                                        'SF_abs', 'SF_abs_sign',
                                                                                        'SF_validation','SF_validation_check']]

                                ## Save to output folder
                                csvFile = SF_temp_file_path_vn + "\\" + f'SF_calc_{flex_pf_id}_{specific_network_element_name}_{reqid}_{resp_id}.csv'
                                if len(csvFile) >= 255:
                                    csvFile = SF_temp_file_path_vn + "\\" + f'SF_calc_{flex_pf_id[:5]}_{specific_network_element_name[-7:]}_{reqid}_{resp_id}.csv'
                                network_element_df_all_steps_2.to_csv(csvFile, index=True)
                            else:
                                asset_loading_initial = 0
                                asset_loading_final = 0
                                sensitivity_factor = const.SF_default
                                sensitivity_factor_direction = const.SF_default
                                specific_network_element_name = None

                            sf_calculated_datetime = datetime.now()
                            print(f"Flex Asset [{flex_pf_id}] dispatched with Pgini: [{offered_power_mw:.2f}]MW has a SF of [{sensitivity_factor:.3f}] wrt [{specific_network_element_name}]")
                            oGen_flex.pgini = 0 # Reset dispatch of power for flex asset for further runs this might need to change to a variable rather than zero for existing contracts
                            print(f'Reset flex asset [{oGen_flex.loc_name}] Pgini to [0]MW')
                            print('='*65)
                            ## Record the loading and sensitivity factor ##
                            offered_power_kw = offered_power_mw * 1e3  # Convert MW to kW
                            dfPSASFResponses = PSA_SF_Responses_table(reqid, resp_id, bsp, primary, feeder, secondary,
                                               terminal, constrained_pf_id, constrained_pf_type, asset_loading_initial,
                                               asset_loading_final, busbar_from, busbar_to,scenario, start_time,
                                               duration,required_power_kw, flex_pf_id, offered_power_kw, calculate_historical,
                                               sf_calculated_datetime, response_datetime,entered_datetime,
                                               sensitivity_factor,sensitivity_factor_direction)
                            dfPSASFResponses.loc[dfPSASFResponses["sensitivity_factor"] <
                                                 float(dict_config[const.strPSASFThreshold]), "sensitivity_factor"] = 0
                            lDfPSASFResponses.append(dfPSASFResponses)

                else:
                    for reqid in setReqids:
                        dfReqid = groupedbyreqid.get_group(reqid)
                        sf_calculated_datetime = datetime.now()
                        sensitivity_factor = const.SF_default
                        asset_loading = 0
                        asset_loading_initial = 0
                        asset_loading_final = 0
                        constrained_pf_id = dfReqid.iloc[0]['constrained_pf_id']
                        constrained_pf_type = dfReqid.iloc[0]['constrained_pf_type']
                        for index, row in dfReqid.iterrows():
                            if bResponse:
                                resp_id = row['resp_id']
                            else:
                                resp_id = row['con_id']
                            bsp = row['bsp']
                            primary = row['primary']
                            feeder = row['feeder']
                            secondary = row['secondary']
                            terminal = row['terminal']
                            required_power_kw = row['required_power_kw']
                            busbar_from = row['busbar_from']
                            busbar_to = row['busbar_to']
                            start_time = row['start_time']
                            duration = row['duration']
                            response_datetime = row['response_datetime']
                            entered_datetime = row['entered_datetime']
                            flex_pf_id = row['flex_pf_id']
                            offered_power_kw = row['offered_power_kw']
                            sensitivity_factor_direction = const.SF_default
                            dfPSASFResponses = PSA_SF_Responses_table(reqid, resp_id, bsp, primary, feeder, secondary,
                                               terminal,constrained_pf_id, constrained_pf_type, asset_loading_initial,
                                               asset_loading_final, busbar_from,busbar_to,scenario,start_time,duration,
                                               required_power_kw,flex_pf_id, offered_power_kw,calculate_historical,
                                               sf_calculated_datetime,response_datetime,entered_datetime,
                                               sensitivity_factor,sensitivity_factor_direction)
                            dfPSASFResponses.loc[dfPSASFResponses["sensitivity_factor"] <
                                                 float(dict_config[const.strPSASFThreshold]), "sensitivity_factor"] = 0
                            lDfPSASFResponses.append(dfPSASFResponses)

    df_final = pd.concat(lDfPSASFResponses)
    df_final = df_final.sort_values(by=['resp_id'])
    df_final = ut.expand_dataframe(dfSNDoutput_copy, df_final)
    #Output results to PSArunID folder
    if bResponse:
        fName = PSArunID + const.strPSAResponses + "-" + vn + ".xlsx"
        df_final.to_excel(os.path.join(PSAfolder,fName), index=False)
    else:
        fName = PSArunID + const.strPSAContracts + "-" + vn + ".xlsx"
        df_final.to_excel(os.path.join(PSAfolder,fName), index=False)

    #Output results to PSA and SND shared folder if required
    if bOutput2SND:
        msg = fName + " " + "has been successfully saved to PSA_SND_SHARED_FOLDER: " + file_detector_folder
        df_final.to_excel(os.path.join(file_detector_folder, fName), index=False)

    print("Sensitivity factors successfully calculated!")
    print('='*65)
    print('\n\n\n\n')
    return status, msg

if __name__ == "__main__":
    args = sys.argv
    # args[0] = module name
    # args[1] = function name
    # args[2:] = function args : (*unpacked)
    globals()[args[1]](*args[2:])