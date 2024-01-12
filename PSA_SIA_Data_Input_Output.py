import os
import re
import pandas as pd
import numpy as np
from datetime import date, datetime, timedelta
from io import BytesIO
import urllib, json
from urllib.request import Request, urlopen
from base64 import b64encode
import json
import requests
import PSA_SND_Constants as const
import PSA_SND_Utilities as ut
import PF_Config as pfconf
print("################################################################")
print('----- PF_Config parameters -----')
print(f'Running in debug mode: [{pfconf.bDebug}]')
print(f'Running with PowerFactory GUI: [{pfconf.bShowPFAPP}]')
print(f'Output loading of ALL assets: [{pfconf.bOutputAllLoading}]')
print(f'Monitoring loading of ALL assets: [{pfconf.bMonitorAllAssets}]')
print(f'Use NeRDA: [{pfconf.bUseNeRDA}]')
print("################################################################")

dict_config = ut.PSA_SND_read_config(const.strConfigFile)
strSIAParams = dict_config[const.strSIAParams].replace("\'","")
lSIAParams = []
for item in strSIAParams.split(","):
    lSIAParams.append(item)
listtest = dict_config[const.strSIAParams]

def SIA_get_token(user, pwd, timeout):
    PSAstatusStr = ""
    token = ""
    PSAstatus = const.PSAok
    # headers = {"accept": "application/json","Content Type": "application/json", }
    headers = {"accept": "application/json", "Content-Type": "application/json", }
    payload = {"username": user, "password": pwd}
    url = f"{dict_config[const.strSIABaseURL]}{dict_config[const.strSIALoginURL]}"
    try:
        response = requests.post(url=url, headers=headers, json=payload, timeout=timeout)
    except requests.exceptions.Timeout:
        PSAstatus = const.PSAfileReadError
        PSAstatusStr = PSAstatusStr + "Session times out when accessing SIA token" + "\n"
    except requests.exceptions.TooManyRedirects:
        PSAstatus = const.PSAfileReadError
        PSAstatusStr = PSAstatusStr + "SIA URL is bad. Please try another one." + "\n"
    except requests.exceptions.RequestException as e:
        raise SystemExit(e)
    else:
        if response.status_code == 400:
            PSAstatus = const.PSAfileReadError
            PSAstatusStr = PSAstatusStr + "Fetching SIA Token ERROR 400: BAD REQUEST for: " + url + "\n"
        elif response.status_code == 403:
            PSAstatus = const.PSAfileReadError
            PSAstatusStr = PSAstatusStr + "Fetching SIA Token ERROR 403: " + url + " is FORBIDDEN" + "\n"
        elif response.status_code == 404:
            PSAstatus = const.PSAfileReadError
            PSAstatusStr = PSAstatusStr + "Fetching SIA Token ERROR 404: " + url + " is NOT FOUND" + "\n"
        elif response.status_code == 500:
            PSAstatus = const.PSAfileReadError
            PSAstatusStr = PSAstatusStr + "Fetching SIA Token ERROR 500: INTERNAL SERVER ERROR - " + url + "\n"
        elif response.status_code == 502:
            PSAstatus = const.PSAfileReadError
            PSAstatusStr = PSAstatusStr + "Fetching SIA Token ERROR 502: BAD GATEWAY - " + url + "\n"
        elif response.status_code == 503:
            PSAstatus = const.PSAfileReadError
            PSAstatusStr = PSAstatusStr + "Fetching SIA Token ERROR 503: SERVICE IS UNAVAILABLE - " + url + "\n"
        if PSAstatus != const.PSAok:
            return token, PSAstatus, PSAstatusStr
        if "Invalid credentials" in response.content.decode('utf-8'):
            PSAstatus = const.PSAfileReadError
            PSAstatusStr = PSAstatusStr + "Invalid SIA credentials" + "\n"
        else:
            token = response.json()['access_token']

    return token, PSAstatus, PSAstatusStr

# Return the day index for the SIA data array
def SIA_dIndex(t1, t2):
    d = t2 - t1
    return int(d.days)

# Return the half hour index for the SIA data array
def SIA_hIndex(t1, t2):
    d = t2 - t1
    return int(d.seconds/60/30)

# Initialise arrays for SIA data
def SIA_initData(number_of_days, number_of_half_hours):
    global SIAdata_gen_mva
    global SIAdata_gen_mw
    global SIAdata_net_dem_mva
    global SIAdata_net_dem_mw
    global SIAdata_und_dem_mva
    global SIAdata_und_dem_mw

    SIAdata_gen_mva, SIAdata_gen_mw, SIAdata_net_dem_mva, SIAdata_net_dem_mw, SIAdata_und_dem_mva, SIAdata_und_dem_mw = \
        (np.zeros((number_of_half_hours,number_of_days)) for i in range(6))
    # dictFeeder,dictGroup,dictGen = ({k: pd.DataFrame() for k in pfconf.SIAParams} for i in range(3))
    dictFeeder, dictGroup, dictGen = ({k: pd.DataFrame() for k in lSIAParams} for i in range(3))

    # for key in dictFeeder.keys:
    #     dictFeeder[key] = pd.DataFrame()
    # for key in dictGroup.keys:
    #     dictGroup[key] = pd.DataFrame()
    # for key in dictGen.keys:
    #     dictGen[key] = pd.DataFrame()

    return dictFeeder, dictGroup, dictGen
# Get security authorization for API

# Read list of feeders and groups of SIA assets from test input file

# Load SIA data for each asset from API into jdata array
def createSIAfolder(PSAfolder):
    # path = os.path.join(PSAfolder, "\\SIA_FOLDER")
    path = PSAfolder + "\\SIA_DATA"
    # path = os.path.join(PSAfolder, PSArunID)
    if not os.path.isdir(path):
        os.mkdir(path)
    return path

def readLineTrafoNameThreshold(assets):
    dictLines = dict()
    dictTrafos = dict()
    for index, row in assets.iterrows():
        if row['ASSET_TYPE'] == const.strAssetLine:
            dictLines[row['ASSET_ID']] = row['MAX_LOADING']
        elif row['ASSET_TYPE'] == const.strAssetTrans:
            dictTrafos[row['ASSET_ID']] = row['MAX_LOADING']
    return dictLines,dictTrafos

def readLineTrafoNames(assets):
    lLines = []
    lTrafos = []
    for index, row in assets.iterrows():
        if row['ASSET_TYPE'] == const.strAssetLine:
            lLines.append(row['ASSET_ID'])
        if row['ASSET_TYPE'] == const.strAssetTrans:
            lTrafos.append(row['ASSET_ID'])
    return lLines, lTrafos

def loadSIAData(strbUTC,assets,timeout,OutputFilePath,token, bDebug, PSArunID, number_of_days, number_of_half_hours):
    Days = pd.Series(['%s%s' % ('Day',d) for d in range(number_of_days)])
    print("################################################################")
    print("Querying SIA API")
    print("################################################################")
    dataOK = True
    isDayZero = True
    PSAstatus = const.PSAok
    PSAstatusStr = "loadSIAData"
    jdata=[]
    dtobj1 = "unknown_start_time"
    start_datetime = "unknown_start_datetime"
    AssetType = ""
    urlstr = ""
    dictSIAFeeder, dictSIAGen, dictSIAGroup = (dict() for i in range(3))
    lFeeders, lGroups, lGens = ([] for i in range(3))
    for index, row in assets.iterrows():
        if row['SIA_ID'] != "NONE":
            hdr = {'Authorization': 'Bearer ' + token}
            prms = {'id':row['SIA_ID'],'format':'json'}
            if 'Feeder' in row['ASSET_TYPE']:
                AssetType = "Feeder"
                lFeeders.append(row['SIA_ID'])
                urlstr = dict_config[const.strSIAURLFeeder]
            elif 'Group' in row['ASSET_TYPE']:
                AssetType = "Group"
                lGroups.append(row['SIA_ID'])
                urlstr = dict_config[const.strSIAURLGroup]
            elif 'Generation' in row['ASSET_TYPE']:
                AssetType = "Generator"
                lGens.append(row['SIA_ID'])
                urlstr = dict_config[const.strSIAURLGenerator]
            else:
                dataOK = False
                PSAstatus = const.PSAfileReadError
                PSAstatusStr = PSAstatusStr + row['ASSET_TYPE'] + " is not a valid ASSET_TYPE in Asset Data" + "\n"
            if dataOK:
                try:
                    resp = requests.get(urlstr, headers=hdr, params=prms, timeout=timeout)
                except requests.exceptions.Timeout:
                    PSAstatus = const.PSAfileReadError
                    PSAstatusStr = PSAstatusStr + "Session times out when accessing SIA token" + "\n"
                except requests.exceptions.TooManyRedirects:
                    PSAstatus = const.PSAfileReadError
                    PSAstatusStr = PSAstatusStr + "SIA URL is bad. Please try another one." + "\n"
                except requests.exceptions.RequestException as e:
                    raise SystemExit(e)
                else:
                    if resp.status_code == 400:
                        dataOK = False
                        PSAstatus = const.PSAfileReadError
                        PSAstatusStr = PSAstatusStr + "Requesting SIA Data ERROR 400: BAD REQUEST for" + urlstr + "\n"
                    elif resp.status_code == 403:
                        dataOK = False
                        PSAstatus = const.PSAfileReadError
                        PSAstatusStr = PSAstatusStr + "Requesting SIA Data ERROR 403: " + urlstr + " is FORBIDDEN" + "\n"
                    elif resp.status_code == 404:
                        dataOK = False
                        PSAstatus = const.PSAfileReadError
                        PSAstatusStr = PSAstatusStr + "Requesting SIA Data ERROR 404: " + urlstr + " is NOT FOUND" + "\n"
                    elif resp.status_code == 500:
                        dataOK = False
                        PSAstatus = const.PSAfileReadError
                        PSAstatusStr = PSAstatusStr + "Requesting SIA Data ERROR 500: INTERNAL SERVER ERROR - " + urlstr + "\n"
                    elif resp.status_code == 502:
                        dataOK = False
                        PSAstatus = const.PSAfileReadError
                        PSAstatusStr = PSAstatusStr + "Requesting SIA Data ERROR 502: BAD GATEWAY - " + urlstr + "\n"
                    elif resp.status_code == 503:
                        dataOK = False
                        PSAstatus = const.PSAfileReadError
                        PSAstatusStr = PSAstatusStr + "ERROR 503: SERVICE IS UNAVAILABLE - " + urlstr + "\n"

        # Loop through each asset (uuid), strip out data and save to asset specific Excel file
                    elif resp.status_code == 200:
                        if "error" in resp.content.decode("utf-8"):
                            dataOK = False
                            PSAstatus = const.PSAfileReadError
                            PSAstatusStr = PSAstatusStr + row['SIA_ID'] + " is not a valid SIA_ID in Asset Data" + "\n"
                        else:
                            data2append = resp.json()
                            data2append['type'] = AssetType
                            jdata.append(data2append)

    if dataOK:
        dictFeeder, dictGroup, dictGen = SIA_initData(number_of_days, number_of_half_hours)
        for uuid in jdata:
            arraySIA = np.full([number_of_half_hours, number_of_days], np.nan)
            strDtLast = uuid['data']['forecast'][-1]['timestamp'][:19]

            fname = uuid['metadata']['uuid']
            a = '-'.join(PSArunID.split("-")[1:4])
            if uuid['type'] == "Feeder":
                if a == uuid['metadata']['last_available_timestamp_prediction_demand_mva'].split(' ')[0]:
                    isDayZero = True
                else:
                    isDayZero = False
            if uuid['type'] == "Generator":
                if a == uuid['metadata']['last_available_timestamp_prediction_gen_mw'].split(' ')[0]:
                    isDayZero = True
                else:
                    isDayZero = False
            if uuid['type'] == "Feeder":
                dictSIAFeeder[fname] = dict()
            if uuid['type'] == "Group":
                dictSIAGroup[fname] = dict()
            if uuid['type'] == "Generator":
                dictSIAGen[fname] = dict()
            if isDayZero:
                SIAdata_gen_mva, SIAdata_gen_mw, SIAdata_net_dem_mva, SIAdata_net_dem_mw, SIAdata_und_dem_mva, \
                SIAdata_und_dem_mw=(np.full([number_of_half_hours, number_of_days], np.nan) for i in
                                    range(6))
            else:
                SIAdata_gen_mva, SIAdata_gen_mw, SIAdata_net_dem_mva, SIAdata_net_dem_mw, SIAdata_und_dem_mva, \
                SIAdata_und_dem_mw = (np.full([number_of_half_hours, number_of_days+1], np.nan) for i in
                                      range(6))
        # for uuid in jdata:
        #     fname = uuid['metadata']['uuid']
            first = True
            for jd in uuid['data']['forecast']:
                if first:
                    start_datetime = jd['timestamp'][:19]
                    dtstr1 = jd['timestamp'][:11] + '00:00'
                    if strbUTC == "TRUE":
                        dtstr1 = datetime.strptime(dtstr1, '%Y-%m-%d %H:%M') - timedelta(hours=1)
                        dtobj1 = dtstr1
                    else:
                        dtobj1 = datetime.strptime(dtstr1, '%Y-%m-%d %H:%M')
                    first = False
                dtstr2 = jd['timestamp'][:16]
                dtobj2 = datetime.strptime(dtstr2, '%Y-%m-%d %H:%M')
                d = SIA_dIndex(dtobj1, dtobj2)
                h = SIA_hIndex(dtobj1, dtobj2)
                if isDayZero:
                    iteration_days = number_of_days - 1
                else:
                    iteration_days = number_of_days

                # dtEndDt = dtobj1 + timedelta(days=iteration_days)  # End time of expected matrix

                if d <= iteration_days and h <= number_of_half_hours-1:
                    SIAdata_gen_mva[h][d] = jd['generation_mva']
                    SIAdata_gen_mw[h][d] = jd['generation_mw']
                    SIAdata_net_dem_mva[h][d] = jd['net_demand_mva']
                    SIAdata_net_dem_mw[h][d] = jd['net_demand_mw']
                    SIAdata_und_dem_mva[h][d] = jd['und_demand_mva']
                    SIAdata_und_dem_mw[h][d] = jd['und_demand_mw']

            # Fill in missing data for a larger matrix
            for d in range(1,iteration_days+1):
                for h in range(number_of_half_hours):
                    if np.isnan(SIAdata_gen_mva[h][d]):
                        SIAdata_gen_mva[h][d] = SIAdata_gen_mva[h][d-1]
                        SIAdata_gen_mw[h][d] = SIAdata_gen_mw[h][d-1]
                        SIAdata_net_dem_mva[h][d] = SIAdata_net_dem_mva[h][d-1]
                        SIAdata_net_dem_mw[h][d] = SIAdata_net_dem_mw[h][d-1]
                        SIAdata_und_dem_mva[h][d] = SIAdata_und_dem_mva[h][d-1]
                        SIAdata_und_dem_mw[h][d] = SIAdata_und_dem_mw[h][d-1]

            if not isDayZero:
                SIAdata_gen_mva = np.delete(SIAdata_gen_mva, 0, axis=1)
                SIAdata_gen_mw = np.delete(SIAdata_gen_mw, 0, axis=1)
                SIAdata_net_dem_mva = np.delete(SIAdata_net_dem_mva, 0, axis=1)
                SIAdata_net_dem_mw = np.delete(SIAdata_net_dem_mw, 0, axis=1)
                SIAdata_und_dem_mva = np.delete(SIAdata_und_dem_mva, 0, axis=1)
                SIAdata_und_dem_mw = np.delete(SIAdata_und_dem_mw, 0, axis=1)

            # lParams = dict_config[const.strSIAParams]
            lSIAdata = [SIAdata_gen_mva, SIAdata_gen_mw, SIAdata_net_dem_mva, SIAdata_net_dem_mw, SIAdata_und_dem_mva,
                        SIAdata_und_dem_mw]

            for i in range(len(lSIAParams)):
                dfSIA = pd.DataFrame(lSIAdata[i])
                if dfSIA.empty:
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
                    print(f"{lSIAParams[i]} of {fname} is empty!")
                    print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")

                if uuid['type'] == "Feeder":
                    dictSIAFeeder[fname][lSIAParams[i]] = dfSIA
                elif uuid['type'] == "Group":
                    dictSIAGroup[fname][lSIAParams[i]] = dfSIA
                elif uuid['type'] == "Generator":
                    dictSIAGen[fname][lSIAParams[i]] = dfSIA
    
                    
    if dataOK:                
        print("################################################################")
        print("SIA data successfully retrieved from API")
        print("################################################################")
    else:
        print("################################################################")
        print("Error retrieving SIA data from API")
        print(PSAstatusStr)
        print("################################################################")

    if dataOK:
        Time = pd.Series(['%s:%s' % (h, m) for h in ([00] + list(range(1, 24))) for m in ('00', '30')])
        Time = Time[:number_of_half_hours]
        for strFeeder, dictFeeder in dictSIAFeeder.items():
            #if pfconf.bDebug:
            print("----------------------------------------------------------------")
            print(f"Outputting SIA data for {strFeeder}")
            with pd.ExcelWriter(OutputFilePath + '/' + 'SIA_Feeder_' + strFeeder + '.xlsx', engine='xlsxwriter') as writer:
                for param, dfSIA in dictFeeder.items():
                    dfSIA2excel = dfSIA
                    dfSIA2excel = dfSIA2excel.set_index(Time, drop=False)
                    dfSIA2excel=dfSIA2excel.set_axis(Days, axis=1)
                    dfSIA2excel.to_excel(writer, sheet_name=param, header=True, index=True)

        for strGen, dictGen in dictSIAGen.items():
            #if pfconf.bDebug:
            print("----------------------------------------------------------------")
            print(f"Outputting SIA data for {strGen}")
            with pd.ExcelWriter(OutputFilePath + '/' + 'SIA_Gen_' + strGen + '.xlsx',
                                engine='xlsxwriter') as writer:
                for param, dfSIA in dictGen.items():
                    dfSIA2excel = dfSIA
                    dfSIA2excel = dfSIA2excel.set_index(Time, drop=False)
                    dfSIA2excel=dfSIA2excel.set_axis(Days, axis=1)
                    dfSIA2excel.to_excel(writer, sheet_name=param, header=True, index=True)

        for strGroup, dictGroup in dictSIAGroup.items():
            #if pfconf.bDebug:
            print("----------------------------------------------------------------")
            print(f"Outputting SIA data for {strGroup}")
            with pd.ExcelWriter(OutputFilePath + '/' + 'SIA_Group_' + strGroup + '.xlsx',
                                engine='xlsxwriter') as writer:
                for param, dfSIA in dictGroup.items():
                    dfSIA2excel = dfSIA
                    dfSIA2excel = dfSIA2excel.set_index(Time, drop=False)
                    dfSIA2excel=dfSIA2excel.set_axis(Days, axis=1)
                    dfSIA2excel.to_excel(writer, sheet_name=param, header=True, index=True)

    return dtobj1, start_datetime, PSAstatus, PSAstatusStr, dictSIAFeeder, dictSIAGen, dictSIAGroup