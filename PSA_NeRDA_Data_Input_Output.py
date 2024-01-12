from urllib.error import URLError, HTTPError, ContentTooShortError
from urllib.request import Request, urlopen
from base64 import b64encode
import json
import pandas as pd
import PSA_SND_Constants as const
import PSA_SND_Utilities as ut
from socket import timeout
import os

dict_config = ut.PSA_SND_read_config(const.strConfigFile)

def createNeRDAfolder(foldernumbering,NeRDAfolder):
    path = NeRDAfolder + "\\" + str(foldernumbering) + "NeRDA_DATA"
    if not os.path.isdir(path):
        os.mkdir(path)
    return path

def loadNeRDAData(assets,output_path,user,pwd):
    PSAstatusStr = "loadNeRDAData"
    PSAstatus = const.PSAok
    dfNeRDA = pd.DataFrame()
    # Load NeRDA data
    userAndPassStr = user + ":" + pwd
    userAndPass = b64encode(bytes(userAndPassStr, 'utf-8')).decode("ascii")
    req = Request(dict_config[const.strNERDAURL])
    req.add_header('Authorization', 'Basic %s' % userAndPass)
    try:
        data= urlopen(req,timeout=int(dict_config[const.strNeRDATimeOut]))
    except ContentTooShortError as error:
        PSAstatus = const.PSAfileReadError
        PSAstatusStr = PSAstatusStr + f"Content Too Short Error: NeRDA data not retrieved because\n{error}\nURL: {dict_config[const.strNERDAURL]}" + "\n"

    except timeout as error:
        PSAstatus = const.PSAfileReadError
        PSAstatusStr = PSAstatusStr + f"Timeout Error: NeRDA data not retrieved because\n{error} after {dict_config[const.strNeRDATimeOut]} secs\nURL: {dict_config[const.strNERDAURL]}" + "\n"

    except HTTPError as error:
        PSAstatus = const.PSAfileReadError
        PSAstatusStr = PSAstatusStr + f"HTTP Error: NeRDA data not retrieved because\n{error}\nURL: {dict_config[const.strNERDAURL]}" + "\n"

    except URLError as error:
        if isinstance(error.reason, timeout):
            PSAstatus = const.PSAfileReadError
            PSAstatusStr = PSAstatusStr + f"Timeout Error: NeRDA data not retrieved because\n{error}\nURL: {dict_config[const.strNERDAURL]}" + "\n"

        else:
            PSAstatus = const.PSAfileReadError
            PSAstatusStr = PSAstatusStr + f"URL Error: NeRDA data not retrieved because\n{error}\nURL: {dict_config[const.strNERDAURL]}" + "\n"

    if PSAstatus == const.PSAok:
        lShortName, lstatus, ltimeStamp = ([] for i in range(3))
        JScontent = json.load(data)
        for item in JScontent['value']:
            if item['shortName'] != None:
                lShortName.append(item['shortName'][:9])
            else:
                tstr = item['aliasName'][:9]
                tstr = tstr.replace("~", "_")
                lShortName.append(tstr)
            ltimeStamp.append(item['timeStamp'])
            if str(item['value']) == "0": 
                lstatus.append("CLOSED")
            elif str(item['value']) == "1": 
                lstatus.append("OPEN")
            else:
                lstatus.append("UNKNOWN")
        
        data = {
            "shortName": lShortName,
            "status": lstatus,
            "timeStamp": ltimeStamp,
        }
        # Output to excel file
        dfNeRDA=pd.DataFrame(data)
        dfNeRDA.to_excel(output_path+'-NeRDA_DATA.xlsx', index=False)        
        
    return PSAstatus, PSAstatusStr, dfNeRDA

