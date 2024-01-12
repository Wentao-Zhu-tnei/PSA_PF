import pandas as pd
from datetime import datetime
import numpy as np
import os
import PSA_SND_Constants as const

def mapServices(strbUTC,start_time, dfEvents, PSAfolder, number_of_days, number_of_half_hours):
    Time = pd.Series(['%s:%s' % (h, m) for h in ([00] + list(range(1, int(number_of_half_hours / 2)))) for m in ('00', '30')])

    lAssetIDs = set(dfEvents.iloc[:]["ASSET_ID"])
    print('-'*65)
    print(f'These assets are part of the EVENTS file: {lAssetIDs}')
    print('-'*65)
    lServiceAssets, lDfServices = ([] for i in range(2))

    dictServices = dict()
    for AssetID in lAssetIDs:
        lServiceAssets.append(AssetID)
        arrayService = np.zeros((number_of_days * number_of_half_hours,), dtype=float)
        for index in dfEvents[dfEvents['ASSET_ID'] == AssetID].index:
            deltaStartService = int((dfEvents.loc[index]["START_SERVICE"] - datetime.strptime(str(start_time),
                                     "%Y-%m-%d %H:%M:%S")).total_seconds() / (60 * 30))
            deltaEndService = int((dfEvents.loc[index]["END_SERVICE"] - datetime.strptime(str(start_time),
                                   "%Y-%m-%d %H:%M:%S")).total_seconds() / (60 * 30))

            if strbUTC == "TRUE":
                deltaStartService = deltaStartService - 2
                deltaEndService = deltaEndService - 2
            if deltaStartService < 0:
                indS = 0
            else:
                indS = deltaStartService
            if deltaEndService < 0:
                indE = 0
            else:
                indE = deltaEndService
            arrayService[indS:indE] = dfEvents.loc[index]["P_DISPATCH_KW"]
            
        dfService = pd.DataFrame(0, index=np.arange(number_of_half_hours), columns=np.arange(number_of_days))
        
        for day in range(number_of_days):
            dfService.iloc[range(number_of_half_hours), day] = arrayService[day *
                           number_of_half_hours: (day + 1) * number_of_half_hours]
            lDfServices.append(dfService)
            dfService.set_index(Time)

        dictServices[AssetID] = dfService
        output_folder = PSAfolder + "\\" + const.strEventFiles
        asset_event_file = output_folder + "\\" + AssetID + ".xlsx"
        if not os.path.isdir(output_folder):
            os.makedirs(output_folder)
        dfService.to_excel(asset_event_file)

    return  dictServices

def updateServices(pf,dictDfServices,i_half_hour, i_day, verbose=False):
    for gen, dfFlexServices in dictDfServices.items():
        oGen = pf.app.GetCalcRelevantObjects(gen + ".ElmSym")[0]
        oGen.pgini = dfFlexServices.iloc[i_half_hour, i_day]/1000.0
        if verbose and int(dfFlexServices.iloc[i_half_hour, i_day]) != 0:
            print(str(oGen.loc_name) + " is dispatching: " + str(oGen.pgini) 
                    + " MW at day: " + str(i_day) + " half_hour: " + str(i_half_hour))

