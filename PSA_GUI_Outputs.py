import os
import shutil
from datetime import datetime
import PSA_SND_Utilities as ut
import PSA_SND_Constants as const


def makeResultsFolder(strPSAfolder, PSArunID):
    # strPSAfolder_dict = ut.PSA_SND_read_config(config_file)
    # strPSAfolder = strPSAfolder_dict[results_folder]
    # Create result folder
    if not os.path.exists(strPSAfolder):
        os.mkdir(strPSAfolder)

    PSAfolder = strPSAfolder + "/" + PSArunID
    # Create new folder
    if not os.path.exists(PSAfolder):
        os.mkdir(PSAfolder)

    else:
        PSAstatus = const.PSAfileWriteError
        PSAstatusStr = "Folder already exists: " + PSArunID

    return PSAfolder

def copy2ResultsFolder(PSAfolder,lSrc):
    if bool(lSrc):
        for src in lSrc:
            fnme = os.path.basename(src)
            dst = PSAfolder + "\\" + fnme
            shutil.copy(src, dst)

def getPSARunningMode(plout,unplout):
    PSA_running_mode = "B"
    if plout and not unplout:
        PSA_running_mode = "M"
    elif unplout and not plout:
        PSA_running_mode = "C"
    elif plout and unplout:
        PSA_running_mode = "MC"
    return PSA_running_mode


def PSA_create_man_results_folders(PSAfolder,plout,unplout):
    """
    create results folders for manual mode
    """

    strResultsFolder = PSA_create_man_results_folder(PSAfolder, plout, unplout)
    strInputFolder = PSAfolder + "\\" + const.strINfolder
    strNeRDAFolder = PSAfolder + "\\" + const.strNeRDAfolder
    strSIAFolder = PSAfolder + "\\" + const.strSIAfolder
    if not os.path.isdir(strInputFolder):
        os.mkdir(strInputFolder)
    if not os.path.isdir(strNeRDAFolder):
        os.mkdir(strNeRDAFolder)
    if not os.path.isdir(strSIAFolder):
        os.mkdir(strSIAFolder)
    return strInputFolder, strResultsFolder, strNeRDAFolder, strSIAFolder

def PSA_create_aut_folders(PSAFolder):
    iNewNumber_input_data, strNewFolder_input_data = PSA_create_numbering_folder(PSAFolder,-1,"INPUT_DATA")
    iNewNumber_SIA, strNewFolder_SIA = PSA_create_numbering_folder(PSAFolder, iNewNumber_input_data, "SIA_DATA")
    iNewNumber_NeRDA, strNewFolder_NeRDA = PSA_create_numbering_folder(PSAFolder, iNewNumber_SIA, "NeRDA_DATA")
    iNewNumber_base, strNewFolder_base = PSA_create_numbering_folder(PSAFolder, iNewNumber_NeRDA, "BASE")
    return strNewFolder_input_data, strNewFolder_SIA, strNewFolder_NeRDA, strNewFolder_base

def PSA_create_aut_folders_new(PSAfolder):
    strNewFolder_input_data = PSAfolder + const.strINfolder
    strNewFolder_SIA = PSAfolder + const.strSIAfolder
    strNewFolder_NeRDA = PSAfolder + const.strNeRDAfolder
    strNewFolder_base = PSAfolder + const.strBfolder
    if not os.path.isdir(strNewFolder_input_data):
        os.mkdir(strNewFolder_input_data)
    if not os.path.isdir(strNewFolder_SIA):
        os.mkdir(strNewFolder_SIA)
    if not os.path.isdir(strNewFolder_NeRDA):
        os.mkdir(strNewFolder_NeRDA)
    if not os.path.isdir(strNewFolder_base):
        os.mkdir(strNewFolder_base)
    return strNewFolder_input_data, strNewFolder_SIA, strNewFolder_NeRDA, strNewFolder_base


def PSA_create_man_results_folder(PSAfolder,plout,unplout):
    if plout and not unplout:
        foldername = const.strMfolder
    elif not plout and unplout:
        foldername = const.strCfolder
    elif not plout and not unplout:
        foldername = const.strBfolder
    else:
        foldername = const.strMCfolder

    results_folder = PSAfolder + "\\" + foldername
    if not os.path.isdir(results_folder):
        os.mkdir(results_folder)

    return results_folder

def PSA_create_numbering_folder(PSAFolder, prev_numb,strFoldername):
    iNewNumber = prev_numb+1
    #strNewFolder = PSAFolder + "\\" + str(iNewNumber) + " - " + strFoldername
    if strFoldername == "MAINT":
        strNewFolder = PSAFolder + const.strMfolder
    elif strFoldername == "CONT":
        strNewFolder = PSAFolder + const.strCfolder
    elif strFoldername == "MAINT_CONT":
        strNewFolder = PSAFolder + const.strMCfolder
        
    if not os.path.isdir(strNewFolder):
        os.mkdir(strNewFolder)

    return iNewNumber, strNewFolder

# def PSA_create_aut_folders(PSAfolder,PSArunID,PSA_running_mode):
#     if PSA_running_mode == "B":
#         strFolder_scenario = "BASE"
#     elif PSA_running_mode == "M":
#         strFolder_scenario = "MAINT"
#     elif PSA_running_mode == "C":
#         strFolder_scenario = "CONT"
#     elif PSA_running_mode == "MC":
#         strFolder_scenario = "MAINT_CONT"
#
#     strAutFolder = PSAfolder+ "\\" + PSArunID + "\\" + strFolder_scenario
#     if not os.path.isdir(strAutFolder):
#         os.mkdir(strAutFolder)
#
#     return strAutFolder

def PSA_create_output(config_file, working_folder, BSPmodel, primary, assetFile, bEvents, eventFile, plout, ploutFile,
                      unplout, unploutFile, bSwitch, switchFile, auto):
    bMsg = False
    now = datetime.now()
    nowStr = now.strftime("%Y-%m-%d-%H-%M")
    if auto:
        PSArunID = "AUT_" + nowStr
    else:
        PSArunID = "MAN_" + nowStr

    lModels = [BSPmodel, primary, assetFile, eventFile, ploutFile, unploutFile, switchFile]
    if bEvents:
        pass
    else:
        lModels.remove(eventFile)

    if plout:
        pass
    else:
        lModels.remove(ploutFile)

    if unplout:
        pass
    else:
        lModels.remove(unploutFile)

    if bSwitch:
        pass
    else:
        lModels.remove(switchFile)

    msg = "Please select input file for: " + "\n" + "\n"
    fnme_BSPmodel = "NULL"
    fnme_primary = "NULL"
    fnme_assetFile = "NULL"
    fnme_eventFile = "NULL"
    fnme_ploutFile = "NULL"
    fnme_unploutFile = "NULL"
    fnme_switchFile = "NULL"

    strPSAfolder_dict = ut.PSA_SND_read_config(config_file)
    strPSAfolder = strPSAfolder_dict[working_folder]
    PSAfolder = strPSAfolder + "\\" + PSArunID

    if not os.path.isfile(BSPmodel):
        lModels.remove(BSPmodel)
        fnme_BSPmodel = BSPmodel.replace(strPSAfolder,'')
    else:
        fnme_BSPmodel = os.path.basename(BSPmodel)

    if not os.path.isfile(primary):
        lModels.remove(primary)
        fnme_primary = primary.replace(strPSAfolder, '')
    else:
        fnme_primary = os.path.basename(primary)

    if not os.path.isfile(assetFile):
        bMsg = True
        msg = msg + "Asset Data" + "\n"
        lModels.remove(assetFile)
    else:
        fnme_assetFile = os.path.basename(assetFile)

    if bEvents:
        if not os.path.isfile(eventFile):
            bMsg = True
            msg = msg + "Events Data" + "\n"
            lModels.remove(eventFile)
        else:
            fnme_eventFile = os.path.basename(eventFile)

    if bSwitch:
        if not os.path.isfile(switchFile):
            bMsg = True
            msg = msg + "Switch Data" + "\n"
            lModels.remove(switchFile)
        else:
            fnme_switchFile = os.path.basename(switchFile)

    if plout:
        if not os.path.isfile(ploutFile):
            bMsg = True
            msg = msg + "Maintenance Data" + "\n"
            lModels.remove(ploutFile)
        else:
            fnme_ploutFile = os.path.basename(ploutFile)

    if unplout:
        if not os.path.isfile(unploutFile):
            bMsg = True
            msg = msg + "Contingency Data" + "\n"
            lModels.remove(unploutFile)
        else:
            fnme_unploutFile = os.path.basename(unploutFile)

    return msg, bMsg, PSArunID, lModels, fnme_BSPmodel, fnme_primary, fnme_assetFile, fnme_eventFile, fnme_ploutFile, fnme_unploutFile, fnme_switchFile