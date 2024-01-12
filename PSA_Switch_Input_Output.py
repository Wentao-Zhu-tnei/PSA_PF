import pandas as pd
from datetime import datetime
import numpy as np
import os
import PSA_SND_Constants as const

def mapSwitchStatus(strbUTC, start_time, dfSwitch, PSAfolder, number_of_days, number_of_half_hours):
    Time = pd.Series(['%s:%s' % (h, m) for h in ([00] + list(range(1, int(number_of_half_hours / 2)))) for m in ('00', '30')])

    lAssetIDs = set(dfSwitch.iloc[:]["ASSET_ID"])
    print('-' * 65)
    print(f'These assets are part of the SWITCH file: {lAssetIDs}')
    print('-' * 65)
    # lSwitches, lDfSwitches, lDfSwitchAction = ([] for i in range(3))
    lSwitches = []
    dictSwitchStatus = {'Switch': {'Asset_Status': {}, 'Asset_Action': {}},
                        'Coupler': {'Asset_Status': {}, 'Asset_Action': {}}}
    for AssetID in lAssetIDs:
        lSwitches.append(AssetID)
        arraySwitch = np.zeros((number_of_days * number_of_half_hours,), dtype=float)
        arraySwitchAction = np.full(number_of_days * number_of_half_hours, False)
        for index in dfSwitch[dfSwitch['ASSET_ID'] == AssetID].index:
            deltaStartSwitchStatus = int((dfSwitch.loc[index]["START_SERVICE"] - datetime.strptime(str(start_time),
                                          "%Y-%m-%d %H:%M:%S")).total_seconds() / (60 * 30))
            deltaEndSwitchStatus = int((dfSwitch.loc[index]["END_SERVICE"] - datetime.strptime(str(start_time),
                                        "%Y-%m-%d %H:%M:%S")).total_seconds() / (60 * 30))

            if strbUTC == "TRUE":
                deltaStartSwitchStatus = deltaStartSwitchStatus - 2
                deltaEndSwitchStatus = deltaEndSwitchStatus - 2
            if deltaStartSwitchStatus < 0:
                indS = 0
            else:
                indS = deltaStartSwitchStatus
            if deltaEndSwitchStatus < 0:
                indE = 0
            else:
                indE = deltaEndSwitchStatus
            bSwitchAction = True
            if dfSwitch.loc[index]["ACTION"] == "CLOSED":
                iSwitchStatus = 1
            else:
                iSwitchStatus = 0
            arraySwitch[indS:indE] = iSwitchStatus
            arraySwitchAction[indS:indE] = bSwitchAction

        dfSwitchStatus = pd.DataFrame(0, index=np.arange(number_of_half_hours),
                                 columns=np.arange(number_of_days))
        dfSwitchAction = pd.DataFrame(False, index=np.arange(number_of_half_hours),
                                 columns=np.arange(number_of_days))

        for day in range(number_of_days):
            dfSwitchStatus.iloc[range(number_of_half_hours), day] = arraySwitch[day * number_of_half_hours:
                                                                                  (day + 1) * number_of_half_hours]
            dfSwitchAction.iloc[range(number_of_half_hours), day] = arraySwitchAction[day * number_of_half_hours:
                                                                                  (day + 1) * number_of_half_hours]
            # lDfSwitches.append(dfSwitchStatus)
            dfSwitchStatus.set_index(Time)
            # lDfSwitchAction.append(dfSwitchAction)
            dfSwitchAction.set_index(Time)

        output_folder = PSAfolder + "\\" + const.strSwitchFiles
        asset_switch_file = output_folder + "\\" + AssetID + ".xlsx"
        if not os.path.isdir(output_folder):
            os.makedirs(output_folder)
        writer = pd.ExcelWriter(asset_switch_file, engine='xlsxwriter')
        dfSwitchStatus.to_excel(writer, sheet_name='Status')
        dfSwitchAction.to_excel(writer, sheet_name='Action')
        writer.save()
        # Assign the asset switch status to the correct dictionary key based on the asset type
        asset_type = dfSwitch.loc[dfSwitch['ASSET_ID'] == AssetID, 'ASSET_TYPE'].iloc[0]
        dictSwitchStatus[asset_type]['Asset_Action'][AssetID] = dfSwitchAction
        dictSwitchStatus[asset_type]['Asset_Status'][AssetID] = dfSwitchStatus

    return dictSwitchStatus

def updateSwitchStatus(pf, dictSwitches, i_half_hour, i_day, verbose=False):
    for switch, dfSwitchAction in dictSwitches['Switch']['Asset_Action'].items():
        oSwitch = pf.app.GetCalcRelevantObjects(switch + ".StaSwitch")[0]
        if dictSwitches['Switch']['Asset_Action'][switch].iloc[i_half_hour, i_day]:
            oSwitch.on_off = dictSwitches['Switch']['Asset_Status'][switch].iloc[i_half_hour, i_day].item()
            if verbose:
                strAction = ('OPENED' if oSwitch.on_off == 0 else 'CLOSED')
                print(str(oSwitch.loc_name) + " is " + strAction + " at day: " + str(i_day) + " half_hour: " + str(i_half_hour))
    for coupler, dfCouplerAction in dictSwitches['Coupler']['Asset_Action'].items():
        oCoupler = pf.app.GetCalcRelevantObjects(coupler + ".ElmCoup")[0]
        if dictSwitches['Coupler']['Asset_Action'][coupler].iloc[i_half_hour, i_day]:
            oCoupler.on_off = dictSwitches['Coupler']['Asset_Status'][coupler].iloc[i_half_hour, i_day].item()
            if verbose:
                strAction = ('OPENED' if oCoupler.on_off == 0 else 'CLOSED')
                print(str(oCoupler.loc_name) + " is " + strAction + " at day: " + str(i_day) + " half_hour: " + str(i_half_hour))

    # for switch, dfSwitchStatuses in dictSwitches['Switch'].items():
    #     oSwitch = pf.app.GetCalcRelevantObjects(switch + ".StaSwitch")[0]
    #     oSwitch.on_off = dfSwitchStatuses.iloc[i_half_hour, i_day].item()
    #     if verbose:
    #         if int(dfSwitchStatuses.iloc[i_half_hour, i_day]) == 0:
    #             print(str(oSwitch.loc_name) + " is OPEN at day: " + str(i_day) + " half_hour: " + str(i_half_hour))
    # for coupler, dfCouplerStatuses in dictSwitches['Coupler'].items():
    #     oCoupler = pf.app.GetCalcRelevantObjects(coupler + ".ElmCoup")[0]
    #     oCoupler.on_off = dfCouplerStatuses.iloc[i_half_hour, i_day].item()
    #     if verbose:
    #         if int(dfCouplerStatuses.iloc[i_half_hour, i_day]) == 0:
    #             print(str(oCoupler.loc_name) + " is OPEN at day: " + str(i_day) + " half_hour: " + str(i_half_hour))


