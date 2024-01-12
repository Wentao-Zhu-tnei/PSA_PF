import os
import pandas as pd
from dateutil import parser
import datetime as dt
from datetime import datetime
import PSA_SND_Constants as const
import PSA_PF_Functions as pffun
import PF_Config as pfconf
import powerfactory_interface as pf_interface
import PSA_SND_Utilities as ut
import PSA_ProgressBar as pbar

def readRefAssetNames(RefAssetFile):
    lRefLine, lRefTrafo, lRefSyncMachine, lRefSwitch, lRefCoupler = ([] for i in range(5))
    dfRefAsset = pd.read_excel(RefAssetFile)
    for index, row in dfRefAsset.iterrows():
        if row['Type'] == const.strAssetGen:
            lRefSyncMachine.append(row['Asset'])
        if row['Type'] == const.strAssetTrans:
            lRefTrafo.append(row['Asset'])
        if row['Type'] == const.strAssetCoupler:
            lRefCoupler.append(row['Asset'])
        if row['Type'] == const.strAssetLine:
            lRefLine.append(row['Asset'])
        if row['Type'] == const.strAssetSwitch:
            lRefSwitch.append(row['Asset'])
    return lRefLine, lRefTrafo, lRefSyncMachine, lRefSwitch, lRefCoupler
def validatePFMaintNames(lRefLine, lRefTrafo, dfMaint):
    status = const.PSAok
    msg = ""
    for idx, row in dfMaint.iterrows():
        if row['ASSET_TYPE'] == const.strAssetLine:
            if row['ASSET_ID'] not in lRefLine:
                status = const.PSAfileReadError
                msg = msg + f"\n {row['ASSET_TYPE']} {row['ASSET_ID']} in Maintenance File not found in PowerFactory model."
        if row['ASSET_TYPE'] == const.strAssetTrans:
            if row['ASSET_ID'] not in lRefTrafo:
                status = const.PSAfileReadError
                msg = msg + f"\n {row['ASSET_TYPE']} {row['ASSET_ID']} in Maintenance File not found in PowerFactory model."
    return status, msg
def validatePFContNames(lRefLine, lRefTrafo, dfCont):
    status = const.PSAok
    msg = ""
    for idx, row in dfCont.iterrows():
        if row['ASSET_TYPE'] == const.strAssetLine:
            if row['ASSET_ID'] not in lRefLine:
                status = const.PSAfileReadError
                msg = msg + f"\n {row['ASSET_TYPE']} {row['ASSET_ID']} in Contingency File not found in PowerFactory model."
        if row['ASSET_TYPE'] == const.strAssetTrans:
            if row['ASSET_ID'] not in lRefTrafo:
                status = const.PSAfileReadError
                msg = msg + f"\n {row['ASSET_TYPE']} {row['ASSET_ID']} in Contingency File not found in PowerFactory model."
    return status, msg
def validatePFEventNames(lRefSyncMachine, dfEvent):
    status = const.PSAok
    msg = ""
    for idx, row in dfEvent.iterrows():
        if row['ASSET_TYPE'] == const.strAssetGen:
            if row['ASSET_ID'] not in lRefSyncMachine:
                status = const.PSAfileReadError
                msg = msg + f"\n {row['ASSET_TYPE']} {row['ASSET_ID']} in Events File not found in PowerFactory model."
    return status,msg
def validatePFAssetNames(lRefLine, lRefTrafo, dfAsset):
    status = const.PSAok
    msg = ""
    for idx, row in dfAsset.iterrows():
        if row['ASSET_TYPE'] == const.strAssetLine:
            if row['ASSET_ID'] not in lRefLine:
                status = const.PSAfileReadError
                msg = msg + f"\n {row['ASSET_TYPE']} {row['ASSET_ID']} in Asset File not found in PowerFactory model."
        if row['ASSET_TYPE'] == const.strAssetTrans:
            if row['ASSET_ID'] not in lRefTrafo:
                status = const.PSAfileReadError
                msg = msg + f"\n {row['ASSET_TYPE']} {row['ASSET_ID']} in Asset File not found in PowerFactory model."
    return status, msg
def validatePFSwitchNames(lRefSwitch, lRefCoupler, dfSwitch):
    # Switch file not yet implemented
    status = const.PSAok
    msg = ""
    for idx, row in dfSwitch.iterrows():
        if row['ASSET_TYPE'] == "Switch":
            if row['ASSET_ID'] not in lRefSwitch:
                status = const.PSAfileReadError
                msg = msg + f"\n {row['ASSET_TYPE']} {row['ASSET_ID']} in Switch File not found in PowerFactory model."
    for idx, row in dfSwitch.iterrows():
        if row['ASSET_TYPE'] == "Coupler":
            if row['ASSET_ID'] not in lRefCoupler:
                status = const.PSAfileReadError
                msg = msg + f"\n {row['ASSET_TYPE']} {row['ASSET_ID']} in Switch File not found in PowerFactory model."
    return status,msg
def validateSNDFileNames(lRefSyncMachine, SNDResponsesFile):
    # Switch file not yet implemented
    status = const.PSAok
    msg = ""
    dfSNDResponses = pd.read_excel(SNDResponsesFile)
    strResponsesFileName = os.path.basename(SNDResponsesFile)
    for idx, row in dfSNDResponses.iterrows():
        if row['flex_pf_id'] not in lRefSyncMachine:
            status = const.PSAfileReadError
            msg = msg + f"\n Flexibility asset {row['flex_pf_id']} in {strResponsesFileName} not found in PowerFactory model."
    return status,msg
def PSA_SND_read_params(PSAFolder,PSArunID):
    dict_params = dict()
    params_file_path = PSAFolder + "\\" + PSArunID + "\\" + PSArunID + const.strRunTimeParams
    para_file = open(params_file_path,'r')
    lines = para_file.readlines()

    for line in lines:
        strPSAfolder = line
        strSplit = strPSAfolder.split()
        key = strSplit[0]
        strSplit.remove(key)
        value = ' '.join(strSplit)
        dict_params[key] = value

    return dict_params

def readAssetFile(assetFile,sheet_name):
    asset_data = pd.read_excel(assetFile, sheet_name=sheet_name)
    return asset_data

def readEventFile(eventFile,sheet_name):
    event_data = pd.read_excel(eventFile, sheet_name=sheet_name, converters={'START_SERVICE': pd.to_datetime,
                                                                             'END_SERVICE': pd.to_datetime})
    return event_data

def readSwitchFile(switchFile,sheet_name):
    switch_data = pd.read_excel(switchFile, sheet_name=sheet_name, converters={'START_SERVICE': pd.to_datetime,
                                                                             'END_SERVICE': pd.to_datetime})
    return switch_data

def readMaintFile(maintFile,sheet_name):
    maint_data = pd.read_excel(maintFile, sheet_name=sheet_name, converters={'START_OUTAGE': pd.to_datetime,
                                                                             'END_OUTAGE': pd.to_datetime})
    return maint_data

def readContFile(contFile,sheet_name):
    cont_data = pd.read_excel(contFile, sheet_name=sheet_name, converters={'START_OUTAGE': pd.to_datetime,
                                                                             'END_OUTAGE': pd.to_datetime})
    return cont_data

def checkDataCol(dfData,lCol):
    # Check asset data
    dataOK = True
    PSAstatus = const.PSAok
    PSAstatusStr = "" + "\n"

    missing_item = [item for item in lCol if item not in dfData.columns]
    extra_item = [item for item in dfData.columns if item not in lCol]

    if bool(missing_item):
        dataOK = False
        PSAstatus = const.PSAfileReadError
        PSAstatusStr = PSAstatusStr + "Columns required: " + str(missing_item) + "\n"

    #if bool(extra_item):
    #    dataOK = False
    #    PSAstatus = const.PSAfileReadError
    #    PSAstatusStr = PSAstatusStr + "Unknown columns found: " + str(extra_item) + "\n"

    return dataOK, PSAstatus, PSAstatusStr

def validateDateTime(datetime):
    msg = "Data is correct"
    error_msg = "Incorrect data format, should be YYYY-MM-DD hh:mm"
    try:
        isValid = bool(parser.parse(str(datetime)))
    except:
        isValid = False
        msg = error_msg

    return isValid, msg

def checkAssetID(df1,df2):
    AssetID1 = df1.loc[:,"ASSET_ID"].values.tolist()
    AssetID2 = df2.loc[:,"ASSET_ID"].values.tolist()

    missing_elements = []
    for item in AssetID2:
        if item not in AssetID1:
            missing_elements.append(item)

    if bool(missing_elements):
        dataOK_AssetID = False
    else:
        dataOK_AssetID = True
    return dataOK_AssetID, missing_elements

def checkOutageData(dfData,typeOutage):

    dataOK = True
    PSAstatus = const.PSAok
    PSAstatusStr = ""

    if (typeOutage == "Events") or (typeOutage == "Switch"):
        strStart = "START_SERVICE"
        strEnd = "END_SERVICE"
    else:
        strStart = "START_OUTAGE"
        strEnd = "END_OUTAGE"

    # Check outage data, does assetID exist in assetReg and are dates/times valid
    for index, row in dfData.iterrows():
        # Are times valid and startOutage before endOutage
        if dataOK:
            isValid_start, msg_start = validateDateTime(row[strStart])
            isValid_end, msg_end = validateDateTime(row[strEnd])

            if not isValid_start:
                dataOK = False
                PSAstatus = const.PSAfileReadError
                PSAstatusStr = PSAstatusStr + msg_start + ": " + str(row[strStart]) + "\n"

            elif not isValid_end:
                dataOK = False
                PSAstatus = const.PSAfileReadError
                PSAstatusStr = PSAstatusStr + msg_end + ": " + str(row[strEnd]) + "\n"

            else:
                try:
                    duration = row[strEnd] - row[strStart]
                except BaseException as err:
                    PSAstatus = const.PSAfileReadError
                    PSAstatusStr = PSAstatusStr + "Error in {} Outage file in row (time): ".format(typeOutage) + str(row['ASSET_ID']) + "\n"
                    # Changed to return message, previously was blank return statement:
                    return dataOK, PSAstatus, PSAstatusStr

                if duration.days < 0:
                    dataOK = False
                    PSAstatus = const.PSAfileReadError
                    PSAstatusStr = PSAstatusStr + "Error in {} Outage file in row (time): ".format(typeOutage) + str(row['ASSET_ID']) + "\n"

    return dataOK, PSAstatus, PSAstatusStr

def checkExcelSheet(excel_file,lTarget_sheets):
    dataOK = True
    lSheetNames = pd.ExcelFile(excel_file).sheet_names
    PSAstatus = const.PSAok
    PSAstatusStr = ""
    if isinstance(lTarget_sheets, list):
        for target_sheet in lTarget_sheets:
            if target_sheet not in lSheetNames:
                dataOK = False
                PSAstatus = const.PSAfileReadError
                PSAstatusStr = "Excel sheet {} not found in {}.\n".format(target_sheet,os.path.basename(excel_file))
    else:
        if lTarget_sheets not in lSheetNames:
            dataOK = False
            PSAstatus = const.PSAfileReadError
            PSAstatusStr = "Excel sheet {} not found in {}.\n".format(lTarget_sheets, os.path.basename(excel_file))
    return dataOK,PSAstatus,PSAstatusStr

def checkOverload(strHHmode,df,threshold,start_date):
    half_hours, days, values = ([] for i in range(3))
    temp1 = df.shape[1]
    temp2 = df.shape[0]
    for i_col in range(temp1):
        for i_row in range(temp2):
            diff = threshold - float(df.iloc[i_row, i_col])
            if diff < 0:
                days.append(i_col)
                half_hours.append(i_row)
                values.append(df.iloc[i_row, i_col])

    A1, B1, C1 = ([[] for _ in range(len(days))] for i in range(3))
    A_temp, B_temp, C_temp, = ([] for i in range(3))
    iList = 0

    if bool(days):
        if strHHmode == "TRUE":
            for i in range(len(days)):
                A_temp.append(days[i])
                B_temp.append(half_hours[i])
                C_temp.append(values[i])
                A1[iList].append(A_temp)
                B1[iList].append(B_temp)
                C1[iList].append(C_temp)
                A_temp = []
                B_temp = []
                C_temp = []
                iList += 1
        else:
            for i in range(len(days)):
                if i == 0:
                    A_temp.append(days[i])
                    B_temp.append(half_hours[i])
                    C_temp.append(values[i])
                    if i == 0 and len(days) == 1:
                        A1[iList].append(A_temp)
                        B1[iList].append(B_temp)
                        C1[iList].append(C_temp)
                elif (days[i] == days[i - 1] and half_hours[i] == half_hours[i - 1] + 1) or (
                        days[i] == days[i - 1] + 1 and half_hours[i] == df.first_valid_index() and half_hours[
                    i - 1] == df.last_valid_index()):
                    if i == len(days) - 1:
                        A_temp.append(days[i])
                        B_temp.append(half_hours[i])
                        C_temp.append(values[i])
                        A1[iList].append(A_temp)
                        B1[iList].append(B_temp)
                        C1[iList].append((C_temp))
                    else:
                        if i != 0:
                            A_temp.append(days[i])
                            B_temp.append(half_hours[i])
                            C_temp.append(values[i])
                else:
                    if i != 0:
                        A1[iList].append(A_temp)
                        B1[iList].append(B_temp)
                        C1[iList].append(C_temp)
                        A_temp, B_temp, C_temp, = ([] for i in range(3))
                        iList += 1
                        A_temp.append(days[i])
                        B_temp.append(half_hours[i])
                        C_temp.append(values[i])
                        if i == len(days) - 1 and len(days) != 1:
                            A1[iList].append(A_temp)
                            B1[iList].append(B_temp)
                            C1[iList].append(C_temp)

        A1 = [item for sublist in A1 for item in sublist]
        B1 = [item for sublist in B1 for item in sublist]
        C1 = [item for sublist in C1 for item in sublist]

        lStartingDay, lStartingHour, lDuration, lAveLoading, lMaxLoading, lMinLoading, index_event, lStart_time, lEnd_time \
            = ([] for i in range(9))

        for i_list in range(len(A1)):
            lStartingDay.append(A1[i_list][0])
            lStartingHour.append(B1[i_list][0])
            lDuration.append(len(C1[i_list]) * 0.5)
            lAveLoading.append(sum(C1[i_list]) / len(C1[i_list]))
            lMaxLoading.append(max(C1[i_list]))
            lMinLoading.append(min(C1[i_list]))

        msg = "There are {} overloading events happened at given threshold {} \n".format(len(lStartingHour), threshold)

        for i_event in range(len(lStartingHour)):
            date_time = datetime.strptime(start_date, "%Y-%m-%d") + dt.timedelta(days=lStartingDay[i_event],
                                                                                 minutes=lStartingHour[i_event] * 30)
            msg = msg + "No. {} event (ave. loading: {}, max. loading {}) started at {}, lasted for {} hours\n". \
                format(i_event + 1, lAveLoading[i_event], lMaxLoading[i_event], str(date_time), lDuration[i_event])
            index_event.append("Event" + str(i_event + 1))
            end_time = date_time + dt.timedelta(minutes=60 * lDuration[i_event])
            lStart_time.append(str(date_time))
            lEnd_time.append(str(end_time))

        # dict_event = {'Start Time': lStart_time, 'End Time': lEnd_time, 'Duration (hr)': lDuration, 'Maximum Loading': lMaxLoading,
        #               'Minimum Loading': lMinLoading, 'Average Loading': lAveLoading}
        lDuration_new = []
        for Duration in lDuration:
            hh = int(Duration*60//60)
            hh = str(0) + str(hh) if len(str(hh)) == 1 else str(hh)
            mm = int(Duration*60%60)
            mm = str(0) + str(mm) if len(str(mm)) == 1 else str(mm)
            lDuration_new.append(f"{hh}:{mm}")
        lThreshold = [threshold for i in range(len(lMaxLoading))]
        dict_event = {'start_time': lStart_time, 'duration': lDuration_new, 'maximum_loading': lMaxLoading,
                      'loading_threshold_pct': lThreshold}
        df_event = pd.DataFrame(dict_event, index=[index_event])
        bEvent = True
    else:
        msg = "There are no overloading events happened at given threshold {} \n".format(threshold)
        # df_event = pd.DataFrame(data={"No overloading events detected."})
        df_event = pd.DataFrame()
        bEvent = False
    return bEvent, msg, df_event

# def checkOverload_hh_mode(df,threshold,start_date):
#     half_hours, days, values = ([] for i in range(3))
#     temp1 = df.shape[1]
#     temp2 = df.shape[0]
#     for i_col in range(temp1):
#         for i_row in range(temp2):
#             diff = threshold - float(df.iloc[i_row, i_col])
#             if diff < 0:
#                 days.append(i_col)
#                 half_hours.append(i_row)
#                 values.append(df.iloc[i_row, i_col])
#
#     A1, B1, C1 = ([[] for _ in range(len(days))] for i in range(3))
#     A_temp, B_temp, C_temp, = ([] for i in range(3))
#     iList = 0
#
#     if bool(days):
#         for i in range(len(days)):
#             A_temp.append(days[i])
#             B_temp.append(half_hours[i])
#             C_temp.append(values[i])
#             A1[iList].append(A_temp)
#             B1[iList].append(B_temp)
#             C1[iList].append(C_temp)
#             A_temp = []
#             B_temp = []
#             C_temp = []
#             iList += 1
#
#         A1 = [item for sublist in A1 for item in sublist]
#         B1 = [item for sublist in B1 for item in sublist]
#         C1 = [item for sublist in C1 for item in sublist]
#
#         lStartingDay, lStartingHour, lDuration, lAveLoading, lMaxLoading, lMinLoading, index_event, lStart_time, lEnd_time \
#             = ([] for i in range(9))
#
#         for i_list in range(len(A1)):
#             lStartingDay.append(A1[i_list][0])
#             lStartingHour.append(B1[i_list][0])
#             lDuration.append(len(C1[i_list]) * 0.5)
#             lAveLoading.append(sum(C1[i_list]) / len(C1[i_list]))
#             lMaxLoading.append(max(C1[i_list]))
#             lMinLoading.append(min(C1[i_list]))
#
#         msg = "There are {} overloading events happened at given threshold {} \n".format(len(lStartingHour), threshold)
#
#         for i_event in range(len(lStartingHour)):
#             date_time = datetime.strptime(start_date, "%Y-%m-%d") + dt.timedelta(days=lStartingDay[i_event],
#                                                                                  minutes=lStartingHour[i_event] * 30)
#             msg = msg + "No. {} event (ave. loading: {}, max. loading {}) started at {}, lasted for {} hours\n". \
#                 format(i_event + 1, lAveLoading[i_event], lMaxLoading[i_event], str(date_time), lDuration[i_event])
#             index_event.append("Event" + str(i_event + 1))
#             end_time = date_time + dt.timedelta(minutes=60 * lDuration[i_event])
#             lStart_time.append(str(date_time))
#             lEnd_time.append(str(end_time))
#
#         # dict_event = {'Start Time': lStart_time, 'End Time': lEnd_time, 'Duration (hr)': lDuration, 'Maximum Loading': lMaxLoading,
#         #               'Minimum Loading': lMinLoading, 'Average Loading': lAveLoading}
#         lDuration_new = []
#         for Duration in lDuration:
#             hh = int(Duration*60//60)
#             hh = str(0) + str(hh) if len(str(hh)) == 1 else str(hh)
#             mm = int(Duration*60%60)
#             mm = str(0) + str(mm) if len(str(mm)) == 1 else str(mm)
#             lDuration_new.append(f"{hh}:{mm}")
#         dict_event = {'start_time': lStart_time, 'duration_hr': lDuration_new, 'maximum_loading': lMaxLoading}
#         df_event = pd.DataFrame(dict_event, index=[index_event])
#         bEvent = True
#     else:
#         msg = "There are no overloading events happened at given threshold {} \n".format(threshold)
#         # df_event = pd.DataFrame(data={"No overloading events detected."})
#         df_event = pd.DataFrame()
#         bEvent = False
#     return bEvent, msg, df_event

def validateAssetType(dfAsset):
    dataOK = True
    PSAstatus = const.PSAok
    PSAstatusStr = ""
    lAssetTypesInput = dfAsset.ASSET_TYPE.to_list()

    lAssetTypes = [const.strAssetGen, const.strAssetLine, const.strAssetLoad, const.strAssetTrans,
                   const.strAssetSwitch, const.strAssetCoupler]

    for strAssetTypeInput in lAssetTypesInput:
        if strAssetTypeInput not in lAssetTypes:
            dataOK = False
            PSAstatus = const.PSAfileReadError
            PSAstatusStr = PSAstatusStr + strAssetTypeInput + " is an unknown asset type" + "\n"

    return dataOK, PSAstatus, PSAstatusStr

def validateFlexReqtsData(fnme_BSPmodel, proj_dir, refAssetFile, assetFile,bEvent,eventFile,plout,ploutFile,unplout,unploutFile, bSwitch, switchFile):
    nameOK_Asset, nameOK_Event, nameOK_Maint, nameOK_Cont = (const.PSAok for i in range(4))
    dataOK_xlSheetAsset = True
    dataOK_xlSheetEvent = True
    dataOK_xlSheetSwitch = True
    dataOK_xlSheetPlout = True
    dataOK_xlSheetUnplout = True
    dataOK_col_AssetData = True
    dataOK_col_EventData = True
    dataOK_col_SwitchData = True
    dataOK_col_MaintData = True
    dataOK_col_ContData = True
    dataOK_EventData = True
    dataOK_SwitchData = True
    dataOK_MaintData = True
    dataOK_ContData = True
    dataOK_AssetDataAssetType = True
    dataOK_EventDataAssetType = True
    dataOK_SwitchDataAssetType = True
    dataOK_ploutDataAssetType = True
    dataOK_unploutDataAssetType = True
    msg_Asset, msg_Event, msg_Maint, msg_Cont, PSAstatusStr_xlSheetEvent, \
    PSAstatusStr_xlSheetPlout, PSAstatusStr_xlSheetUnplout = ("" for i in range(7))
    # Update asset reference file
    if pfconf.bCreateRef:
        dict_config = ut.PSA_SND_read_config(const.strConfigFile)
        # Name of new asset ref data file
        assetRefFile = const.strRefAssetFile[:-5] + "_" + fnme_BSPmodel[:-4] + ".xlsx"
        fullAssetRefFile = os.path.join(dict_config[const.strWorkingFolder], assetRefFile)
        if not os.path.exists(fullAssetRefFile):
            pbar.createProgressBar("Updating Asset Data Reference from Power Factory", False)
            pb=5
            pbar.updateProgressBar(pb, "Starting Power Factory")
            print("Updating asset reference data for PowerFactory Assets....")
            pf_object, status, msg = pf_interface.run(dict_config[const.strPFUser], pfconf.bMultiProcess)
            pb=33
            pbar.updateProgressBar(pb, "Loading PF Model")
            pf_interface.activate_project(pf_object, fnme_BSPmodel, proj_dir)
            pb=66
            pbar.updateProgressBar(pb, "Retrieving Assets")
            pffun.PF_getMasterRefList(pf_object, fullAssetRefFile)
            pb=100
            pbar.updateProgressBar(pb, "Finished")
            pf_interface.PF_exit(pf_object)
            pbar.delProgressBar()
            print("Reference list of PowerFactory Assets has been updated!")

    # Read in asset reference file
    lRefLine, lRefTrafo, lRefSyncMachine, lRefSwitch, lRefCoupler = readRefAssetNames(refAssetFile)
    # Check input data column
    isAssetFile = os.path.isfile(assetFile)
    if isAssetFile:
        [dataOK_xlSheetAsset, PSAstatus_xlSheetAsset, PSAstatusStr_xlSheetAsset] = \
            checkExcelSheet(assetFile, [const.strSIASheet, const.strAssetSheet, const.strNeRDASheet])
        if dataOK_xlSheetAsset:
            dfAssetData = readAssetFile(assetFile,const.strAssetSheet)
            nameOK_Asset, msg_Asset = validatePFAssetNames(lRefLine, lRefTrafo, dfAssetData)
            if nameOK_Asset == const.PSAok:
                dataOK_AssetDataAssetType, PSAstatus_AssetDataAssetType, PSAstatusStr_AssetDataAssetType = validateAssetType(
                    dfAssetData)
                if dataOK_AssetDataAssetType:
                    [dataOK_col_AssetData, PSAstatus_col_AssetData, PSAstatusStr_col_AssetData] = checkDataCol(dfAssetData,
                                                                                                      const.lColAssetData)

    if bEvent:
        [dataOK_xlSheetEvent, PSAstatus_xlSheetEvent, PSAstatusStr_xlSheetEvent] = checkExcelSheet(eventFile,
                                                                                                   const.strDataSheet)

        if dataOK_xlSheetEvent:
            dfEventData = readEventFile(eventFile,const.strDataSheet)
            nameOK_Event, msg_Event = validatePFEventNames(lRefSyncMachine, dfEventData)
            if nameOK_Event == const.PSAok:
                [dataOK_col_EventData, PSAstatus_col_EventData, PSAstatusStr_col_EventData] = checkDataCol(dfEventData,
                                                                                                           const.lColEventData)
                if dataOK_col_EventData:
                    dataOK_EventDataAssetType, PSAstatus_EventDataAssetType, PSAstatusStr_EventDataAssetType = validateAssetType(dfEventData)
                    if dataOK_EventDataAssetType:
                        [dataOK_EventData, PSAstatus_EventData, PSAstatusStr_EventData] = checkOutageData(dfEventData,"Events")

    if bSwitch:
        [dataOK_xlSheetSwitch, PSAstatus_xlSheetSwitch, PSAstatusStr_xlSheetSwitch] = checkExcelSheet(switchFile,
                                                                                                   const.strDataSheet)

        if dataOK_xlSheetSwitch:
            dfSwitchData = readSwitchFile(switchFile,const.strDataSheet)
            nameOK_Switch, msg_Switch = validatePFSwitchNames(lRefSwitch, lRefCoupler, dfSwitchData)
            if nameOK_Switch == const.PSAok:
                [dataOK_col_SwitchData, PSAstatus_col_SwitchData, PSAstatusStr_col_SwitchData] = checkDataCol(dfSwitchData,
                                                                                                           const.lColSwitchData)
                if dataOK_col_SwitchData:
                    dataOK_SwitchDataAssetType, PSAstatus_SwitchDataAssetType, PSAstatusStr_SwitchDataAssetType = validateAssetType(dfSwitchData)
                    if dataOK_SwitchDataAssetType:
                        [dataOK_SwitchData, PSAstatus_SwitchData, PSAstatusStr_SwitchData] = checkOutageData(dfSwitchData,"Switch")


    if plout:
        [dataOK_xlSheetPlout, PSAstatus_xlSheetPlout, PSAstatusStr_xlSheetPlout] = checkExcelSheet(ploutFile,
                                                                                                   const.strDataSheet)
        if dataOK_xlSheetPlout:
            dfMaintData = readMaintFile(ploutFile,const.strDataSheet)
            nameOK_Maint, msg_Maint = validatePFMaintNames(lRefLine, lRefTrafo, dfMaintData)
            if nameOK_Maint == const.PSAok:
                [dataOK_col_MaintData, PSAstatus_col_MaintData, PSAstatusStr_col_MaintData] = checkDataCol(dfMaintData,
                                                                                                           const.lColOutageData)
                if dataOK_col_MaintData:
                    dataOK_ploutDataAssetType, PSAstatus_ploutDataAssetType, PSAstatusStr_ploutDataAssetType = validateAssetType(
                        dfMaintData)
                    if dataOK_ploutDataAssetType:
                        [dataOK_MaintData, PSAstatus_MaintData, PSAstatusStr_MaintData] = checkOutageData(dfMaintData,"Maintenance")

    if unplout:
        [dataOK_xlSheetUnplout, PSAstatus_xlSheetUnplout, PSAstatusStr_xlSheetUnplout] = checkExcelSheet(unploutFile,
                                                                                                   const.strDataSheet)
        if dataOK_xlSheetUnplout:
            dfContData = readContFile(unploutFile,const.strDataSheet)
            nameOK_Cont, msg_Cont = validatePFContNames(lRefLine, lRefTrafo, dfContData)
            if nameOK_Cont == const.PSAok:
                [dataOK_col_ContData, PSAstatus_col_ContData, PSAstatusStr_col_ContData] = checkDataCol(dfContData,
                                                                                                        const.lColOutageData)
                if dataOK_col_ContData:
                    dataOK_unploutDataAssetType, PSAstatus_unploutDataAssetType, PSAstatusStr_unploutDataAssetType = validateAssetType(
                        dfContData)
                    if dataOK_unploutDataAssetType:
                        [dataOK_ContData, PSAstatus_ContData, PSAstatusStr_ContData] = checkOutageData(dfContData, "Contingency")

    # if isAssetFile and plout and dataOK_xlSheetAsset and dataOK_xlSheetPlout:
    #     #dataOK_AssetID_plout, missing_IDs_plout = checkAssetID(assetFile,ploutFile)
    #     dataOK_AssetID_plout, missing_IDs_plout = checkAssetID(dfAssetData, dfMaintData)
    # if isAssetFile and unplout and dataOK_xlSheetAsset and dataOK_xlSheetUnplout:
    #     #dataOK_AssetID_unplout, missing_IDs_unplout = checkAssetID(assetFile,unploutFile)
    #     dataOK_AssetID_unplout, missing_IDs_unplout = checkAssetID(dfAssetData, dfContData)
    msg_ok = "All files are validated."
    msg_warning = ""
    bWarning = False
    if nameOK_Asset != const.PSAok:
        bWarning = True
        msg_warning = msg_warning + str(msg_Asset) + "\n"
    if nameOK_Event != const.PSAok:
        bWarning = True
        msg_warning = msg_warning + str(msg_Event) + "\n"
    if nameOK_Maint != const.PSAok:
        bWarning = True
        msg_warning = msg_warning + str(msg_Maint) + "\n"
    if nameOK_Cont != const.PSAok:
        bWarning = True
        msg_warning = msg_warning + str(msg_Cont) + "\n"

    if not dataOK_xlSheetAsset:
        bWarning = True
        msg_warning = msg_warning + str(PSAstatusStr_xlSheetAsset) + "\n"

    if not dataOK_xlSheetEvent:
        bWarning = True
        msg_warning = msg_warning + str(PSAstatusStr_xlSheetEvent) + "\n"

    if not dataOK_xlSheetSwitch:
        bWarning = True
        msg_warning = msg_warning + str(PSAstatusStr_xlSheetSwitch) + "\n"

    if not dataOK_xlSheetPlout:
        bWarning = True
        msg_warning = msg_warning + str(PSAstatusStr_xlSheetPlout) + "\n"

    if not dataOK_xlSheetUnplout:
        bWarning = True
        msg_warning = msg_warning + str(PSAstatusStr_xlSheetUnplout) + "\n"

    # if not dataOK_AssetID_plout:
    #     bWarning = False  # TODO: If this check is useful
    #     if bool(missing_IDs_plout):
    #         msg_warning = msg_warning + "ASSET_IDs between Asset Data File and Maintenance File do not match : " + "\n" + str(missing_IDs_plout) + "\n"
    #
    # if not dataOK_AssetID_unplout:
    #     bWarning = False  # TODO: If this check is useful
    #     if bool(missing_IDs_unplout):
    #         msg_warning = msg_warning + "ASSET_IDs between Asset Data File and Contingency File do not match : " + "\n" + str(missing_IDs_unplout) + "\n"

    if not dataOK_AssetDataAssetType:
        bWarning = True
        msg_warning = msg_warning + "Error found in: " + os.path.basename(assetFile) + "\n" + PSAstatusStr_AssetDataAssetType + "\n"

    if not dataOK_col_AssetData:
        bWarning = True
        msg_warning = msg_warning + "Error found in: " + os.path.basename(assetFile) + "\n" + PSAstatusStr_col_AssetData + "\n"

    if not dataOK_col_EventData:
        bWarning = True
        msg_warning = msg_warning + "Error found in: " + os.path.basename(eventFile) + "\n" + PSAstatusStr_col_EventData + "\n"

    if not dataOK_EventDataAssetType:
        bWarning = True
        msg_warning = msg_warning + "Error found in: " + os.path.basename(eventFile) + "\n" + PSAstatusStr_EventDataAssetType + "\n"

    if not dataOK_col_SwitchData:
        bWarning = True
        msg_warning = msg_warning + "Error found in: " + os.path.basename(switchFile) + "\n" + PSAstatusStr_col_SwitchData + "\n"

    if not dataOK_SwitchDataAssetType:
        bWarning = True
        msg_warning = msg_warning + "Error found in: " + os.path.basename(switchFile) + "\n" + PSAstatusStr_SwitchDataAssetType + "\n"

    if not dataOK_col_MaintData:
        bWarning = True
        msg_warning = msg_warning + "Error found in: " + os.path.basename(ploutFile) + "\n" + PSAstatusStr_col_MaintData + "\n"

    if not dataOK_col_ContData:
        bWarning = True
        msg_warning = msg_warning + "Error found in: " + os.path.basename(unploutFile) + "\n" + PSAstatusStr_col_ContData + "\n"

    if not dataOK_EventData:
        bWarning = True
        msg_warning = msg_warning + "Error found in: " + os.path.basename(eventFile) + PSAstatusStr_EventData + "\n"

    if not dataOK_SwitchData:
        bWarning = True
        msg_warning = msg_warning + "Error found in: " + os.path.basename(switchFile) + PSAstatusStr_SwitchData + "\n"

    if not dataOK_MaintData:
        bWarning = True
        msg_warning = msg_warning + "Error found in: " + os.path.basename(ploutFile) + PSAstatusStr_MaintData + "\n"

    if not dataOK_ploutDataAssetType:
        bWarning = True
        msg_warning = msg_warning + "Error found in: " + os.path.basename(ploutFile) + PSAstatusStr_ploutDataAssetType + "\n"

    if not dataOK_unploutDataAssetType:
        bWarning = True
        msg_warning = msg_warning + "Error found in: " + os.path.basename(unploutFile) + PSAstatusStr_unploutDataAssetType + "\n"

    if not dataOK_ContData:
        bWarning = True
        msg_warning = msg_warning + "Error found in: " + os.path.basename(unploutFile) + PSAstatusStr_ContData + "\n"

    if bWarning == True:
        msg = msg_warning
    else:
        msg = msg_ok

    return msg, bWarning