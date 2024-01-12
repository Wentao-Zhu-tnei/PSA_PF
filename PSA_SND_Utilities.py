"""
PSA_SND utility functions
"""
import os
import pathlib
import pandas as pd
import PSA_SND_Constants as const
import platform
import PF_Config as config
import openpyxl

def open_xl(f_label, root_dir):
    """
    open an excel file from working directory

    input:
    f_label - label where file name is input in GUI
    root_dir - working directory
    """
    full_path_xl = root_dir + f_label.get("1.0", "end").strip()
    if os.path.isfile(full_path_xl):
        os.system('"' + full_path_xl + '"')
    else:
        pass

def out_xl(dfs, sheet_names, outpath, engine):
    writer = pd.ExcelWriter(outpath, engine=engine)
    for i in range(len(dfs)):
        dfs[i].to_excel(writer, sheet_name=sheet_names[i])
    writer.close()

def PSA_SND_read_runtimeparams(filepath):
    rmode=False
    dict_runtimeparams = dict()
    runtimeParamFile = open(filepath, 'r')
    lines = runtimeParamFile.readlines()
    for line in lines:
        strPSAfolder = line
        strPSAfolder = strPSAfolder.replace("\\", "/")
        strSplit = strPSAfolder.split()
        key = strSplit[0]
        strSplit.remove(key)
        value = ' '.join(strSplit)
        dict_runtimeparams[key] = value
        if key == "RUNNING_MODE":
            rmode=True

    if not rmode:
        if dict_runtimeparams["AUTOMATIC"] == "True":
            dict_runtimeparams["RUNNING_MODE"] = const.strRunModeAUT
        else:
            dict_runtimeparams["RUNNING_MODE"] = const.strRunModeMAN

    return dict_runtimeparams

def PSA_SND_read_config(config_file):
    dict_config = dict()
    config_file_path = pathlib.Path(config_file).absolute()
    config_file = open(config_file_path, 'r')
    lines = config_file.readlines()

    for line in lines:
        strPSAfolder = line
        strPSAfolder = strPSAfolder.replace("\\", "/")
        strSplit = strPSAfolder.split()
        key = strSplit[0]
        strSplit.remove(key)
        value = ' '.join(strSplit)
        dict_config[key] = value

    ### Update variables based on hostname and pre-prod platform
    hname = platform.uname()[1]
    print("Hostname = " + hname)
    if hname[:6] == "AZUWVD":
        print("****** CONFIG UPDATE START *****")
        username = os.getlogin()
        dict_config[const.strPFUser] = username
        print("Changed PSA_SND_PF_USER to: " + username)
        #dict_config[const.strWorkingFolder] = "C:\\PSA"
        #print("Changed PSA_SND_ROOT_WORKING_FOLDER to: " + dict_config[const.strWorkingFolder])
        #dict_config[const.strFileDetectorFolder] = "Z:\\PSA_SND_Shared_Folder"
        #print("Changed PSA_SND_FILE_DETECTOR_FOLDER to: " + dict_config[const.strFileDetectorFolder])
        print("****** CONFIG UPDATE END *****")
    return dict_config


def create_dictPSAparams(BSPmodel, primary, assetFile, bEvents, eventFile, plout, ploutFile, unplout, unploutFile, 
                         bSwitch, switchFile, timeStep, kWstep, powFactor,
                         lagging, lne_thresh, trafo_thresh, auto, runMode, numDays):
    dict_PSA_params = dict()

    dict_PSA_params["BSP_MODEL"] = BSPmodel
    dict_PSA_params["PRIMARY"] = primary
    dict_PSA_params["ASSET_DATA"] = assetFile

    dict_PSA_params["EVENTS"] = str(bEvents) 
    if bEvents:
        dict_PSA_params["EVENT_FILE"] = eventFile
    else:
        dict_PSA_params["EVENT_FILE"] = "NULL"
    
    dict_PSA_params["MAINTENANCE"] = str(plout)
    if plout:
        dict_PSA_params["MAINT_FILE"] = ploutFile
    else:
        dict_PSA_params["MAINT_FILE"] = "NULL"

    dict_PSA_params["CONTINGENCY"] = str(unplout)
    if unplout:
        dict_PSA_params["CONT_FILE"] = unploutFile
    else:
        dict_PSA_params["CONT_FILE"] = "NULL"

    dict_PSA_params["SWITCHING"] = str(bSwitch)
    if bSwitch:
        dict_PSA_params["SWITCH_FILE"] = switchFile
    else:
        dict_PSA_params["SWITCH_FILE"] = "NULL"

    dict_PSA_params["TIME_STEP"] = str(timeStep)
    dict_PSA_params["KW_STEP"] = str(kWstep)
    dict_PSA_params["POWER_FACTOR"] = str(powFactor)
    dict_PSA_params["LAGGING"] = str(lagging)
    dict_PSA_params["LINE_THRESHOLD"] = str(lne_thresh)
    dict_PSA_params["TX_THRESHOLD"] = str(trafo_thresh)
    dict_PSA_params["AUTOMATIC"] = str(auto)   
    dict_PSA_params["RUNNING_MODE"] = str(runMode)   
    dict_PSA_params["DAYS"] = str(int(numDays))
    dict_PSA_params["HALF_HOURS"] = str(config.number_of_half_hours)

    return dict_PSA_params


def writePSAparams(PSAfolder, PSArunID, dict_PSA_params):
    dfile = PSAfolder + "\\" + PSArunID + const.strRunTimeParams

    with open(dfile, 'w') as f:
        for key, value in dict_PSA_params.items():
            f.write('%s \t%s\n' % (key, value))

    f.close()

    return


def readPSAparams(PSAparamsFile):
    """
    This function reads the .txt and returns a dictionary with PSA parameters
    """
    dict_PSA_params = dict()

    with open(PSAparamsFile, 'r') as f:
        lines = f.readlines()

    for line in lines:
        strSplit = line.split()
        key = strSplit[0]
        strSplit.remove(key)
        value = ' '.join(strSplit)
        dict_PSA_params[key] = value

    return dict_PSA_params

def readOutageFile(scenario, PSAfolder):
    if scenario == "MAINT":
        strScenario = "M"
    elif scenario == "CONT":
        strScenario = "C"
    elif scenario == "MAINT_CONT":
        strScenario = "MC"
    else:
        strScenario = ""

    #outage_folder = os.path.join(PSAfolder, const.strINfolder, strScenario + "-ASSET_OUTAGE_TABLES")
    ##### join FAILED on longer paths, so replaced with this
    outage_folder = PSAfolder + "\\" + const.strINfolder + "\\" + strScenario + "-ASSET_OUTAGE_TABLES"
    
    # Create two lists for trafo/line status
    dictDfTrafos_2w, dictDfLine, dictDfSynm, dictDfLoad = (dict() for i in range(4))
    for asset_type in (os.listdir(outage_folder)):
        if asset_type == const.strAssetTrans:
            for r, d, f in os.walk(os.path.join(outage_folder, asset_type)):
                for trafo_status_file in f:
                    dfTrafos_2w = pd.read_excel(os.path.join(r, trafo_status_file), index_col=0)
                    dictDfTrafos_2w[trafo_status_file.replace(".xlsx", "")] = dfTrafos_2w
        elif asset_type == const.strAssetLine:
            for r, d, f in os.walk(os.path.join(outage_folder, asset_type)):
                for line_status_file in f:
                    dfLine = pd.read_excel(os.path.join(r, line_status_file), index_col=0)
                    dictDfLine[line_status_file.replace(".xlsx", "")] = dfLine
        elif asset_type == const.strAssetGen:
            for r, d, f in os.walk(os.path.join(outage_folder, asset_type)):
                for synm_status_file in f:
                    dfSynm = pd.read_excel(os.path.join(r, synm_status_file), index_col=0)
                    dictDfSynm[synm_status_file.replace(".xlsx", "")] = dfSynm
        elif asset_type == const.strAssetLoad:
            for r, d, f in os.walk(os.path.join(outage_folder, asset_type)):
                for load_status_file in f:
                    dfLoad = pd.read_excel(os.path.join(r, load_status_file), index_col=0)
                    dictDfLoad[load_status_file.replace(".xlsx", "")] = dfLoad
    return dictDfTrafos_2w, dictDfLine, dictDfSynm, dictDfLoad

def expand_dataframe(dfA, dfB):
    # count the number of duplicates in A
    A_counts = dfA.groupby(['req_id', 'constrained_pf_id', 'flex_pf_id']).size().reset_index(name='count')

    # merge A_counts with B to get the SF column
    B_with_SF = pd.merge(dfB, A_counts, on=['req_id', 'constrained_pf_id', 'flex_pf_id'])

    # repeat each row in B based on the count of duplicates in A
    B_expanded = B_with_SF.loc[B_with_SF.index.repeat(B_with_SF['count'])].reset_index(drop=True).drop(columns=['count'])

    # transformed = pd.merge(B_expanded, dfA[['req_id', 'resp_id', 'constrained_pf_id', 'flex_pf_id']], on=['req_id', 'resp_id'], how='left')

    df_C = pd.merge(dfA, B_expanded, on=['req_id', 'constrained_pf_id', 'flex_pf_id'])

    # Find columns with the same content
    duplicates = df_C.T.duplicated(keep='first')
    duplicates_new = merge_true_false_values(duplicates)
    # duplicates['sensitivity_factor_dir'] = False
    # duplicates['entered_datetime_x'] = False
    # duplicates['secondary_x'] = False
    # duplicates['resp_id_x'] = False
    # duplicates['feeder_x'] = False
    # duplicates['busbar_to_x'] = False
    # duplicates['calculate_historical_x'] = False
    # duplicates['accepted'] = True
    # Drop unnecessary columns and rename remaining columns
    merged_df = df_C.loc[:, ~duplicates_new].rename(
        columns=lambda x: x[:-2] if x.endswith('_x') else x[:-2] if x.endswith('_y') else x)

    merged_df = merged_df.drop_duplicates()
    merged_df = merged_df.reindex(columns=["req_id",	"resp_id",	"bsp",	"primary",	"feeder",	"secondary",
                                           "terminal",	"constrained_pf_id",	"constrained_pf_type", "asset_loading_initial",
                                           "asset_loading_final","busbar_from",	"busbar_to", "scenario", "start_time",
                                           "duration", "required_power_kw",	"flex_pf_id",	"offered_power_kw",
                                           "calculate_historical","sf_calculated_datetime",	"response_datetime",
                                           "entered_datetime","sensitivity_factor",	"sensitivity_factor_dir"])
    return merged_df

def merge_true_false_values(s: pd.Series) -> pd.Series:
    result = pd.Series(index=s.index, dtype=bool)
    for key in s.index:
        if key.endswith("_x"):
            # check if the corresponding key with "_y" exists and its value is different
            key_y = key[:-2] + "_y"
            if key_y in s.index:
                # if so, assign the value "True" to the key with the "True" value
                if s.loc[key]:
                    result[key] = True
                    result[key_y] = False
                else:
                    result[key] = False
                    result[key_y] = True

        elif not key.endswith("_x") and not key.endswith("_y"):
            result[key] = False

    return result

def aggregateSNDresponses(dfSNDresponses):
    dfSNDresponses['resp_id'] = \
        dfSNDresponses.groupby(
            ['req_id', 'accepted', 'bsp', 'primary', 'secondary', 'terminal', 'feeder', 'busbar_from',
             'busbar_to', 'constrained_pf_id', 'constrained_pf_type', 'scenario', 'start_time', 'duration',
             'required_power_kw', 'flex_pf_id', 'response_datetime', 'entered_datetime', 'calculate_historical'])[
            'resp_id'].transform('first')
    return dfSNDresponses.groupby(
        ['req_id', 'resp_id', 'accepted',	'bsp', 'primary', 'secondary', 'terminal', 'feeder', 'busbar_from',
         'busbar_to', 'constrained_pf_id', 'constrained_pf_type',	'scenario', 'start_time',
         'duration', 'required_power_kw', 'flex_pf_id', 'response_datetime', 'entered_datetime', 'calculate_historical']
    )['offered_power_kw'].sum().reset_index()