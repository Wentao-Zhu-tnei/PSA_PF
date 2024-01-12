import pandas as pd
import PSA_SND_Constants as const
import PSA_SND_Utilities as ut
import numpy as np
from datetime import datetime, timedelta

# These functions were originally built for use in a GUI. They have had the 'self' parameter removed.

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Create SND files based on the flex_reqts file and asset_mapping file. It is recommended you truncate the flex_reqts file before testing'''

def create_snd(flex_reqts_file_location, asset_mapping_file_location, relationship, calc_historical, file_type='RESPONSES', out_to_snd=False, truncate=False, accepted=False):
    status = const.PSAok
    msg = 'Creating SND file'
    snd_file_location = 'create_snd'

    dict_config = ut.PSA_SND_read_config(const.strConfigFile)
    topfldr = dict_config[const.strFileDetectorFolder]     
    PSARunID = flex_reqts_file_location.split('/')[-1].removesuffix(f"{const.strPSAFlexReqts}.xlsx")

    df = pd.read_excel(flex_reqts_file_location)
        
    if truncate:
        print('Trunacting')
        df = truncate_flex_reqts_file(df)

    columns = ['req_id','resp_id','bsp','primary','feeder','secondary','terminal','constrained_pf_id','constrained_pf_type',\
               'busbar_from','busbar_to','scenario','start_time','duration','required_power_kw']
        
    for row in list(df.index):
        df.loc[row, 'response_datetime'] = datetime.strptime(df.loc[row, 'start_time'], '%Y-%m-%d %H:%M:%S') - timedelta(days = 2)
    df['entered_datetime'] = df['response_datetime']
    df['calculate_historical'] = calc_historical
        
    try:
        print(asset_mapping_file_location)
        df_map = pd.read_excel(asset_mapping_file_location, usecols=['asset_id','flex_id','flex_power'])
    except:
        status = const.PSAfileReadError
        msg = 'Error reading asset_mapping file'
        return status, msg, snd_file_location
        
    df_pivot = df_map.pivot(index='asset_id', columns='flex_id', values='flex_power')
    
    df_map.rename(columns = {'asset_id': 'constrained_pf_id',
                            'flex_id': 'flex_pf_id',
                            'flex_power': 'offered_power_kw'} ,inplace=True)
    
    map_asset_list = list(set(df_map['constrained_pf_id']))
    df_asset_list = list(set(df['constrained_pf_id']))
    
    if len(map_asset_list) <= len(df_asset_list):
        asset_list = map_asset_list
    else:
        asset_list = df_asset_list
    
    missing_assets = []
    for asset in df_asset_list:
        if asset not in map_asset_list:
            missing_assets.append(asset)
    
    if missing_assets:
        status = const.PSAfileReadError
        msg = f'No flex_id specified for asset_id = {", ".join(missing_assets)}'
        raise ValueError(msg)
        
    if relationship == '1:1':
        df['flex_pf_id'] = np.nan
        df['offered_power_kw'] = np.nan
        for asset in asset_list:
            df_index = df[df['constrained_pf_id']==asset].index
            value = df_pivot.loc[asset].dropna()[0]
            flex_asset = (df_pivot.loc[asset] == value).idxmax()
            df_pivot[flex_asset].loc[asset] = np.nan 

            df.loc[df_index, 'flex_pf_id'] = flex_asset
            df.loc[df_index, 'offered_power_kw'] = value
                
    elif relationship == '1:N':
        df_old = []
        for asset in asset_list:
            df_asset = df[df['constrained_pf_id']==asset]
            df_map_asset = df_map[df_map['constrained_pf_id']==asset]
            req_id = list(df_asset['req_id'])*len(df_map_asset)
            df_new = pd.merge(df_asset, df_map_asset, on='constrained_pf_id',how='inner')
            df_new['req_id']=req_id
            try:
                df_old = pd.concat([df_new, df_old], ignore_index=True)
            except:
                df_old = df_new.copy()
        df = df_old.copy()
        
    df['resp_id'] = df.index
    df['calculate_historical'] = calc_historical 
        
    columns = columns + ['flex_pf_id','offered_power_kw','response_datetime','entered_datetime','calculate_historical']
    df = df.loc[:,columns]
        
    df.insert(2, 'accepted', [accepted]*len(df))
    df.set_index('req_id', inplace=True)

    if file_type == 'RESPONSES':
        snd_file_location = f'{flex_reqts_file_location[:-len(const.strPSAFlexReqts)-5]}{const.strSNDResponses}-V0-0.xlsx'
        if out_to_snd:
            shared_folder = f'{topfldr}/{PSARunID}{const.strSNDResponses}.xlsx'
            df.to_excel(shared_folder)
    else:
        df.rename(columns={'resp_id': 'con_id'}, inplace=True)
        snd_file_location = f'{flex_reqts_file_location.removesuffix(f"{const.strPSAFlexReqts}.xlsx")}{const.strSNDContracts}-V0-0.xlsx'
        if out_to_snd:
            shared_folder = f'{topfldr}/{PSARunID}{const.strSNDContracts}.xlsx'
            df.to_excel(shared_folder)

    df.to_excel(snd_file_location)
    print(f'Saved to {snd_file_location}')
    msg = 'Done'
    return status, msg, snd_file_location
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function takes a sf file location and creates a candidates responses/contracts file'''
def create_candidates(psa_sf_file_location, calc_historical, file_type='RESPONSES', accepted=False):
    status = const.PSAok
    msg = 'Creating SND file'
    snd_file_location = 'create_candidates'

    # Using numbers allows for differences between con_id and resp_id in column 1, meaning both will be accepted
    columns = ['bsp','primary','secondary','terminal','feeder',\
        'busbar_from','busbar_to','constrained_pf_id','constrained_pf_type',\
        'scenario','start_time','duration','required_power_kw','flex_pf_id',\
        'offered_power_kw','response_datetime','entered_datetime']
    '''
    columns = [0,1,2,3,4,5,6,\
            7,8,9,10,\
            11,12,13,14,15,\
            16,19,20]
    '''
    #print(psa_sf_file_location)
    df = pd.read_excel(psa_sf_file_location)
    col_one = list(df)[1]
    df = df.loc[:,['req_id',col_one]+columns]
    df['calculate_historical'] = calc_historical
    df.insert(2, 'accepted', [accepted]*len(df))
    
    fldr = psa_sf_file_location.split("/")
    PSARunID = ('_').join(fldr[-1].split('_')[:2])

    if file_type == 'RESPONSES':
        df.rename(columns={'con_id': 'resp_id'}, inplace=True)
        file = f'{PSARunID}{const.strSNDCandidateResponses}-V0-0.xlsx'
    else:
        df.rename(columns={'resp_id': 'con_id'}, inplace=True)
        file = f'{PSARunID}{const.strSNDCandidateContracts}-V0-0.xlsx'

    df.set_index('req_id', inplace=True)

    snd_file_location = ('/').join(fldr[:-1] + [file])
    df.to_excel(snd_file_location)
    print(f'Saved to {snd_file_location}')
    msg = 'Done'
    return status, msg, snd_file_location

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function takes three values for each asset and for each scenario that had the highest absolute required_power_kw'''
def truncate_flex_reqts_file(df, v=3):
    for scenario in ['BASE','MAINT','CONT','MAINT_CONT']:
        df_scenario = df[df['scenario']==scenario]
        if df_scenario.shape[0] == 0:
            continue
        assets = list(set(df_scenario['constrained_pf_id']))
        for asset in assets:
            df1 = df_scenario[df_scenario['constrained_pf_id']==asset]
            df1 = df1.sort_values('required_power_kw', ascending=False, ignore_index=True, key=lambda values: [abs(value) for value in values])
            df1 = df1.iloc[:v]
            try:
                df2 = pd.concat([df1,df2], ignore_index=True)
            except:
                df2 = df1
    df2.set_index('req_id', inplace=True, drop=False)
    df2.sort_index(inplace=True)
    return df2