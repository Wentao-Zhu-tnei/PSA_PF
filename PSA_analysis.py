# Dependencies line

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
import os
import numpy as np
import PSA_SND_Constants as const
import PSA_SND_Utilities as ut
import PSA_File_Validation as fv
import zipfile
import time

#from matplotlib.ticker import AutoMinorLocator
from matplotlib.patches import Patch
from matplotlib.lines import Line2D
from datetime import datetime, timedelta

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''
Current Known Bugs:

'''

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Functions'''
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''There was a help me function but it was so out of date TB got rid of it.'''

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Returns the filepath that all subsequent functions should you use when locating data. This should be a folder with all the various identities in.'''

def default_directory(PSARunID, directory=None):
    if directory == None:
        config_dict = ut.PSA_SND_read_config(const.strConfigFile)
        directory = f'{config_dict[const.strWorkingFolder]}{const.strDataResults}{PSARunID}'
    elif PSARunID not in directory:
        directory = os.path.join(directory, PSARunID)
        
    return directory

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function auto expands the columns of a pandas dataframes to the maximum width of a variable within the columns. It saves within the same function'''
def auto_expand_and_save(df, final_dir, index=True, sheet_name='Sheet1'):
    if isinstance(df, pd.DataFrame):
        data_df = df
    else:
        data_df = df.data

    writer = pd.ExcelWriter(final_dir, engine='xlsxwriter')
    df.to_excel(writer, sheet_name=sheet_name, na_rep='NaN', index=index)
    worksheet = writer.sheets[sheet_name]

    if index:
        index_columns = data_df.index.names
        i = len(index_columns)
        for idx, col in enumerate(index_columns):
            if len(index_columns) > 1:
                series = data_df.index.levels[idx]
            else:
                series = data_df.index
            max_len = max(series.astype(str).map(len).max(),
                            len(str(col))) * 1.2

            worksheet.set_column(idx, idx, max_len)
    else:
        i = 0

    for idx, col in enumerate(data_df):  # loop through all columns
        idx += i
        series = data_df[col]
        max_len = max(series.astype(str).map(len).max(),  # len of largest item
                        len(str(series.name))) * 1.05

        worksheet.set_column(idx, idx, max_len)  # set column width

    writer.close()
    return

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''A function that saves a result .xlsx if file is a dataframe or .png if else.'''

def save_result(file, file_name, directory, index=True, auto_expand=True, new_folder=True):
    final_dir = ''

    if new_folder:
        directory += const.strAnalysisFiles
        if not os.path.isdir(directory):
            os.mkdir(directory)
    else:
        directory += '\\'

    try:
        styler_test = isinstance(file, pd.io.formats.style.Styler)
    except:
        styler_test = False

    if styler_test or isinstance(file, pd.DataFrame):
        final_dir = f'{directory}{file_name}.xlsx'
        if auto_expand:
            auto_expand_and_save(file, final_dir, index)
        else:
            file.to_excel(final_dir, index=index)

    else:
        final_dir = f'{directory}{file_name}.png'
        plt.savefig(final_dir, dpi=300, bbox_inches='tight', transparent=False)
            
    if final_dir:
        print(f'Saved to {final_dir}')
        return final_dir
    else:
        print('File type could not be inferred')
        return None

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Read the run_time_parameters.txt file within a certain PSARunID and returns the result as a dictionary. This functions used to be called read_config'''

def read_rtp(PSARunID, directory=None):
    directory = default_directory(PSARunID, directory)
    rtpFile = os.path.join(directory, f'{PSARunID}{const.strRunTimeParams}')
    config_dict = ut.PSA_SND_read_runtimeparams(rtpFile)
    # returns the rtp file in the form of a dictionary
    return config_dict

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Essentially an os.join.path that takes by default the default directory and adds a certain target file name, Seldom Used now by still required by some functions'''

def find_data(PSARunID, target, directory=None, scenario='BASE'):
    directory = default_directory(PSARunID, directory=directory)
    
    target = '_'.join(target.split(' ')).upper()
    scenario = '_'.join(scenario.split(' ')).upper()
    
    scenario_dict = {'BASE': const.strBfolder,
                     'MAINT': const.strMfolder,
                     'CONT': const.strCfolder,
                     'MAINT_CONT': const.strMCfolder}

    target_data = f'{directory}{scenario_dict[scenario]}{target}'
    
    if not os.path.isdir(target_data):
        raise FileNotFoundError (f'{target_data} Does not exist!')

    # returns an os filepath of the target folder
    return target_data

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''a function that checks to see if an asset has a specific set threshold in asset_data, returns the config data if not.'''

def get_asset_threshold(asset, PSARunID, directory=None):   
    directory = default_directory(PSARunID, directory=directory)
    config_dict = read_rtp(PSARunID, directory)
    asset_data = pd.read_excel(f'{directory}\\{const.strINfolder}\\{config_dict["ASSET_DATA"]}', index_col='ASSET_ID')
    
    try:
        return float(asset_data.at[asset, 'MAX_LOADING'])
    except:
        if asset[-1].isnumeric(): # Check to see if transformer or line
            return float(config_dict['LINE_THRESHOLD'])
        else:
            return float(config_dict['TX_THRESHOLD'])

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''a function that uses the date part of a folders PSARunID to convert day0 - day10 to their respective days and dates. Note this only returns a list, it then needs to be applied to the appropriate dataframe'''

def convert_to_dates(PSARunID, df_line):
    try:
        date = PSARunID.split('_')[1]
    except:
        date = '-'.join(PSARunID.split('-')[1:6])
    start_date = datetime.strptime(date, '%Y-%m-%d-%H-%M')
    dates = []

    # Make a list of each column 
    for col in list(df_line):
        if type(col) == int:
            dt_object = start_date + timedelta(days=col)
        elif type(col) == str:
            dt_object = start_date + timedelta(days=int(col[3:]))
        # Column names should either be in the format str 'day0' 'day1'... etc or int 0, 1, 2... etc
            
        dates.append(dt_object.strftime('%A')[:3] + ' ' + dt_object.strftime('%d/%m/%y'))

    # Make a dict of old names: new names
    column_names = dict(zip(list(df_line), dates))

    return column_names

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''a function that converts a value to zero if it's nan'''

def nan_to_zero(value):
    return np.nan if value==0 else value


# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''a function that takes a string of one or two words and returns one or two initials, e.g maint_cont becomes MC, maint becomes M'''

def get_initials(scenario):
    initial_1 = scenario[0]
    
    if len(scenario) == 2:
        initial_2 = scenario[1]
    else:
        try: 
            scenario = (scenario.split('_'))
            initial_2 = scenario[1][0]
        except:
            initial_2 = ''
            
    initials = initial_1 + initial_2
    
    return initials

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''a function that takes a value from a dataframe and sets the background colour based on whether the input is positive or negative. It is used by Overview'''
def highlight_red_blue(value):
    colour = '#E77471' if value > 0 else '#77BFC7'
    return f'background-color: {colour}'

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''a function that makes a dataframe of every constrained asset in a certain PSARunID'''

def overview(PSARunID, scenario=None, directory=None, single='', save=False):  
    directory = default_directory(PSARunID, directory=directory)

    if scenario == None: 
        scenario = [fldr for fldr in [const.strBfolder, const.strMfolder, const.strMCfolder, const.strCfolder] if os.path.isdir(f'directory\\{fldr}')][0]
    else:
        scenario = '_'.join(scenario.split(' ')).upper()
    
    flex_reqts = find_data(PSARunID, 'FLEX_REQTS', directory=directory, scenario=scenario)
    df_flex = pd.read_excel(os.path.join(flex_reqts, f'{PSARunID}{const.strPSAFlexReqts}.xlsx'))

    df_flex['start_time'] = pd.to_datetime(df_flex['start_time'])
    df_flex['date'] = df_flex['start_time'].dt.day
    dt = df_flex['start_time'][0].hour + df_flex['start_time'][0].minute / 60

    asset_list = []
    adr = [] # average daily requirement
    daily_max = []
    acp = [] # average constrained periods
    tcd = [] # total contstrained days
    asset_type = []
    asset_threshold = []
    scp = [] # sum of daily requirements
    
    # For each constrained asset, find its:
    # No. of days with at least one constraint
    # Average number of HH Constrained (on days with at least one constraint)
    # Average required power to resolve the constraint across all days
    # Global maximum power requried to resolve the constraint across all days

    unique_assets = list(set(df_flex['constrained_pf_id']))
    unique_assets.sort()
    print(f'Calculating overview for... {len(unique_assets)} Assets')

    for asset in unique_assets:
        df_asset = df_flex.where(df_flex['constrained_pf_id'] == asset).dropna()
        Days = len(set(df_asset['date']))
        Type = df_asset['constrained_pf_type'].iloc[0]
        if Type == const.strPFassetLine:
            Type = const.strAssetLine
        elif Type == const.strPFassetTrans:
            Type = const.strAssetTrans
        elif Type == const.strPFassetGen:
            Type = const.strAssetGen
        elif Type == const.strPFassetLoad:
            Type = const.strAssetLoad
        else:
            Type = "Unknown"
        asset_list.append(asset)
        scp.append(round(df_asset['required_power_kw'].sum(),3)/1000)
        adr.append(round(df_asset['required_power_kw'].mean(),3))
        if abs(df_asset['required_power_kw'].max()) > abs(df_asset['required_power_kw'].min()):
            daily_max.append(round(df_asset['required_power_kw'].max(),3))
        else:
            daily_max.append(round(df_asset['required_power_kw'].min(),3)) 
        acp.append(int(df_asset.shape[0] / Days))
        tcd.append(Days)
        asset_type.append(Type)
        asset_threshold.append(get_asset_threshold(asset, PSARunID, directory=directory))

    # Output into a dictionary
    print('Done')
    overview_dict = {'Asset Name': asset_list,
                     'Type': asset_type,
                     'Max Loading Threshold': asset_threshold,
                     'Days W/ Constraint': tcd,
                     'Avg No. HH Constrained per Day': acp,
                     'Avg Req (kW)': adr,
                     'Global Max Req (kW)': daily_max,
                     'Sum of Req (MWh)': scp}
                 
# Convert that dictionary into a dataframe
    df_overview = pd.DataFrame.from_dict(overview_dict)
    df_overview['ABS'] = abs(df_overview['Sum of Req (MWh)'])
    df_overview = df_overview.sort_values(by = 'ABS', ignore_index = True, ascending = False)
    df_overview.drop(labels='ABS', axis=1, inplace=True)
    styler = df_overview.style.applymap(highlight_red_blue, subset = pd.IndexSlice[:, ['Avg Req (kW)', 'Global Max Req (kW)', 'Sum of Req (MWh)']])
    styler = styler.set_caption(f'Constrained Assets in {PSARunID}: {scenario.upper()}')
    styler = styler.format(formatter= {'Avg Req (kW)': '{:.3f}', 
                                       'Global Max Req (kW)': '{:.3f}',
                                       'Sum of Req (MWh)': '{:.3f}'})

    si = get_initials(scenario)

    if save:
        return save_result(styler, f'{si}_Overview of {PSARunID}', directory), df_overview
    else:
        return styler

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''returns a list, in the format [BASE, MAINT, CONT, MAINT_CONT] containing values for how many unique assets are constrained over the run.'''

def HL_overview(PSARunID, directory=None):
    directory = default_directory(PSARunID, directory=directory)

    default_dict = {'BASE': "Not Run",
                    'MAINT': "Not Run",
                    'CONT': "Not Run",
                    'MAINT_CONT': "Not Run"}
        
    if f'{PSARunID}{const.strPSAFlexReqts}.xlsx' in os.listdir(directory):
        df_flex = pd.read_excel(os.path.join(directory, f'{PSARunID}{const.strPSAFlexReqts}.xlsx'), usecols=['constrained_pf_id','scenario'])
        # Use cols: constrained_pf_id, scenario
        for scenario in list(set(df_flex['scenario'])):
            df_scenario = df_flex[df_flex['scenario'] == scenario]
            default_dict[scenario] = tuple(sorted(df_scenario['constrained_pf_id'].drop_duplicates()))

    else:    
        print(f'{PSARunID}{const.strPSAFlexReqts}.xlsx not in folder {PSARunID}')
        
    return default_dict

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Returns an asset's type or a dictionary containing assets: asset_types using a particular run's asset data file.'''
def get_asset_type(PSARunID, directory, asset=None):
    directory = default_directory(PSARunID, directory)
    rtp_dict = ut.PSA_SND_read_runtimeparams(f'{directory}\\{PSARunID}{const.strRunTimeParams}')
    asset_data_file = f'{directory}{const.strINfolder}{rtp_dict[const.strAssetData]}'
    df = pd.read_excel(asset_data_file, usecols=['ASSET_ID','ASSET_TYPE'], index_col=0)
    asset_dict = dict(zip(df.index, df['ASSET_TYPE']))
    if asset:
        return asset_dict[asset]
    else:
        return asset_dict
    
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''returns a dataframe, stylised or otherwise, containing no. of unique assets constrained for a multitude of file identities.'''

def VHL_overview(stopping_index=10, directory=None, target='both', decorate=False, single=False, sorted_dir=None, save=False):
    if directory == None:
        # This should be redirected to 'results' file, or could always treat directory as required parameter
        directory = os.path.join(os.getcwd(), 'Data Analysis')
    
    if target == 'both':
        target = ['AUT', 'MAN']
    else:
        target = [target]
    
    column_names = ['BASE','MAINT','CONT','MAINT_CONT']
    VHL_dict = {}

    if single:
        identity = directory[-20:]
        directory = directory[:-20]
        VHL_dict[identity] = [len(x) if type(x) == tuple else x for x in HL_overview(identity, directory=directory).values()]
    else:
        if not sorted_dir:
            # Filtered_dir assumes the first 3 characters are either AUT or MAN, otherwise they won't get included.
            filtered_dir = [identity for identity in os.listdir(directory) if identity[0:3] in target and len(identity)==20]
            # Sorted_dir assumes the PSARunIDs that make it through the filter are in the format XXX_YY-MM-DD-hh-mm such that the 4th - 20th character (zero indexed) contains the date
            sorted_dir = sorted(filtered_dir, key = lambda t: datetime.strptime(t[4:], '%Y-%m-%d-%H-%M'), reverse=True)

        for identity in sorted_dir[:stopping_index]:
            print(f'Calculating for... {identity}')
            if not zipfile.is_zipfile(os.path.join(directory, identity)):
                VHL_dict[identity] = [len(x) if type(x) == tuple else x for x in HL_overview(identity, directory=os.path.join(directory, identity)).values()]
            else:
                print('ZIP FILE')
                continue
        
    df_VHL_overview = pd.DataFrame.from_dict(VHL_dict, orient='index', columns=column_names)

    if decorate:
        df_VHL_overview = df_VHL_overview.style.set_caption('Total No. of Assets with Flex Requirements')\
                                               .set_properties(**{'text-align': 'center'})

    if save:
        return save_result(df_VHL_overview, 'High_Level_Overview', directory, new_folder=False), VHL_dict
    else:
        return df_VHL_overview

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function takes a dictionary in the exact format output by VHL_overview when save=True and converts it to a format to populate a dropdown list. It is used in Event Manager'''

def convert_to_dropdown(dictionary):
    sc = {0: 'B', 1: 'M', 2: 'C', 3: 'MC'}
    dropdown = []
    for key, values in dictionary.items():
        dropdown.append(f'{key} | {(" ").join([f"{sc[i]}X" if type(value) == str else f"{sc[i]}{value}" for i, value in enumerate(values)])}')
    return tuple(dropdown)


# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Returns descriptive statistics on an assets loading using df_line'''

def describe_loading(asset, PSARunID, directory=None, scenario='BASE'):
    df = read_asset(asset, PSARunID, directory=directory, scenario=scenario, plot=False, decorate=False, save=False)
    df_melt = pd.melt(df, var_name='Day', value_name='% Loading').describe()
    return df_melt

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''returns a dataframe, stylised or not, of an asset over the days'''

def read_asset(asset, asset_type, PSARunID, scenario='BASE', directory=None, decorate=True, plot=True, save=False):
    directory = default_directory(PSARunID, directory=directory)
    scenario = '_'.join(scenario.split(' ')).upper()
    
    initials = get_initials(scenario)

    scenario_dict = {'BASE': const.strBfolder,
                     'MAINT': const.strMfolder,
                     'CONT': const.strCfolder,
                     'MAINT_CONT': const.strMCfolder}

    if asset_type == const.strAssetLine or asset_type == 'LN_LDG_DATA':
        file = f'{directory}{scenario_dict[scenario]}LN_LDG_DATA\\PSA_{initials}_Loading-{asset}.xlsx'
    elif asset_type == const.strAssetTrans or asset_type == 'TX_LDG_DATA':
        file = f'{directory}{scenario_dict[scenario]}TX_LDG_DATA\\PSA_{initials}_Loading-{asset}.xlsx'
    
    # Read the excel file
    df_line = pd.read_excel(file, index_col=0)

    # Define vmin and vmax
    vmin = get_asset_threshold(asset, PSARunID, directory=directory)
    vmax = max(df_line.max())

    # If vmin very close to vmax, spread the colours out a bit to an imaginary vmax so it looks better
    if vmin*1.25 > vmax:
        vmax = vmin*1.25

    column_names = convert_to_dates(PSARunID, df_line)
    
    if plot:
        df_line.rename(columns = column_names, inplace = True)
        df_melt = pd.melt(df_line, var_name='Day', value_name='%_Loading')

        title = f'{PSARunID}: {asset}'
        ax = plot_generic_graph(PSARunID, df_melt, title=title, y_label='% Loading', decorate='type_1', h_line=vmin, describer=True)

        if save:
            return save_result(ax, f'{initials}_{asset}_LOADING_PCT_GRAPH', directory)
        else:
            return ax
    else:
        df_line.rename(columns = column_names, inplace = True)

        if decorate:
            # Apply a gradient to constrained values
            df_line = df_line.style.format(precision=3) # This won't affect output to excel
            df_line = df_line.background_gradient(cmap='YlOrRd', low=0.3, text_color_threshold=0.35, vmin = vmin, vmax = vmax)

            # Apply a black colour to 0 values (asset out of service)
            # Apply a green colour to non-constrained values (This overwrites black if no 0 values)
            # Apply a white colour to NaN values (time outside calculation window)
            # Highlight local maxima a purple colour
            # Add a caption to the top left of the dataframe to give context

            df_line = df_line.highlight_max(color = 'purple')\
                            .highlight_max(props = 'color: white')\
                            .highlight_between(color = '#463E3F', left = 0, right = 0, inclusive='both')\
                            .highlight_between(color = '#98BF64', left=0, right=vmin, inclusive='right')\
                            .highlight_null('white')\
                            .set_caption(f'{scenario} | {asset} - Threshold: {str(vmin)}')
                            
        if save:
            return save_result(df_line, f'{initials}_{asset}_HEATMAP', directory, auto_expand=False)
        else:
            return df_line

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''A function that returns a dataframe of the b_loading, merged with the temp conres files to compare current'''
def read_psa_b_loading(PSARunID, directory, asset, asset_type, scenario='BASE', version_number='latest', save=False):
    directory = default_directory(PSARunID, directory)

    scenario_dict = {'BASE': const.strBfolder,
                     'MAINT': const.strMfolder,
                     'MAINT_CONT': const.strMCfolder,
                     'CONT': const.strCfolder}

    if asset_type == const.strAssetTrans:
        file_dir = f'{directory}{scenario_dict[scenario]}PSA_B_TRANSFORMER_LOADING'
    elif asset_type == const.strAssetLine:
        file_dir = f'{directory}{scenario_dict[scenario]}PSA_B_LINE_LOADING'
    else:
        raise KeyError(f'Invalid asset type {asset_type}')
    
    if os.path.isdir(file_dir):
        raise FileNotFoundError(f'{file_dir} Does not exist!')

    if version_number == 'latest':
        fldr = f'{directory}\\TEMP_CR_FILES'
        files = batch_latest_version(fldr, PSARunID, directory=directory)
    else:
        files = f'{directory}\\TEMP_CR_FILES\\V{version_number}'

    conres = [file for file in os.listdir(files) if asset in file]
    if conres == []:
        raise KeyError(f'No conres data for asset {asset}')

    df = pd.DataFrame(columns=['req_id','resp_id','constrained_pf_id'], index=range(len(conres)))
    df_old = []
    for i, file in enumerate(conres):
        file_trim = file[11:-4].split('_')

        df.loc[i] = [file_trim[-2], file_trim[-1], f'{("_").join(file_trim[:-2])}']
        df.loc[i+1] = [file_trim[-2], file_trim[-1], f'{("_").join(file_trim[:-2])}']

        df_new = pd.read_csv(os.path.join(files, file), nrows=1, usecols=['Unnamed: 0',
                                                                     'I_HV',
                                                                     'I_LV',
                                                                     'reqID_day_hh'])
        df_new['resp_id'] = file_trim[-1]
        df_new['req_id']  = file_trim[-2]
        df_new['constrained_pf_id'] = f'{("_").join(file_trim[:-2])}'
        try:
            df_old = pd.concat([df_new, df_old], ignore_index=True)
        except:
            df_old = df_new.copy()

    days_list = df_old['reqID_day_hh'].to_list()
    days_list = [f"{file_dir}\\{day.split('_')[2]}d-{day.split('_')[4]}hr.xlsx" for day in days_list]

    lv_values, hv_values = [], []
    for idx, file in enumerate(days_list):
        df_b = pd.read_excel(file, usecols=['Name','bus1/bushlv current','bus2/bushv current'], index_col=0)
        lv_values.append(df_b.loc[asset, 'bus1/bushlv current'])
        hv_values.append(df_b.loc[asset, 'bus2/bushv current'])

    df_old = df_old.rename(columns={'Unnamed: 0':'calc_step'})
    df_old['bushlv_current'] = lv_values
    df_old['bushv_current'] = hv_values
    df_old['scenario'] = scenario
    
    #df_old.set_index(['scenario','constrained_pf_id','req_id','resp_id'])
    df_old = df_old[['constrained_pf_id','calc_step','reqID_day_hh','req_id','resp_id','I_HV','bushv_current','I_LV','bushlv_current']]

    if save:
        return save_result(df_old, f'{asset}_PSA_B_Loading', directory)
    else:
        return df_old
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''a function that finds all the performance logs for all the scenarios within a particular PSARunID'''

def find_logs(PSARunID, directory=None, scenarios=None):
    files_dict = {}
    possible_scenarios = ['BASE', 'MAINT', 'CONT', 'MAINT_CONT'] 
    
    directory = default_directory(PSARunID, directory=directory)

    if scenarios == None: 
        scenarios = possible_scenarios
    elif type(scenarios) == list: 
        scenarios = ['_'.join(scenario.split(' ')).upper() for scenario in scenarios]
    else: 
        scenarios = ['_'.join(scenarios.split(' ')).upper()]
    
    # Convert scenario names to their associated folder name
    scenario_dict = {'BASE': const.strBfolder, 
                     'MAINT': const.strMfolder,
                     'CONT': const.strCfolder,
                     'MAINT_CONT': const.strMCfolder}
                    
    for scenario in scenarios:
        #print(f'{directory}{scenario_dict[scenario][:-1]}') #For Debugging
        if os.path.isdir(f'{directory}{scenario_dict[scenario][:-1]}'):
            initials = get_initials(scenario)
            files_dict[f'{directory}{scenario_dict[scenario][:-1]}'] = f'{initials}_performance_log.csv'
    
    # returns a dictionary with os filepaths as keys and peformance_log filenames as values
    return files_dict

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Converts a particular aspect of the performance log from being plot across the whole time period, to having columns for each day'''
# This function could be replaced by a df.pivot() and is kind of unnecessary, especially as its dependents are no longer used
'''
def split_days(df_log, string):
    # df: the dataframe that contains the column with name string
    # string: a string representing the column name of the performance log data you want to plot
    daily_dict = {}
    nan_list = []

    while len(nan_list) < (48 - df_log[df_log['DAY'] == 0].shape[0]):
        nan_list.append(np.nan)

    for day in list(set(df_log['DAY'])):
        if day == 0:
            values_list = list(df_log[df_log['DAY'] == day][string])
            daily_dict[f'Day0'] = nan_list + values_list
        else:
            daily_dict[f'Day{day}'] = list(df_log[df_log['DAY'] == day][string])
    
    df_daily = pd.DataFrame.from_dict(daily_dict)

    return df_daily
'''
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Plots all or some of the performance logs as as a matplotlib graph'''
# This function has been replaced by plot_multiple_pl
'''
def plot_pl(string, PSARunID, directory=None, scenarios=None, save=False):
    directory = default_directory(PSARunID, directory=directory)
    
    string = '_'.join(string.split(' ')).upper()
    files_dict = find_logs(PSARunID, directory=directory, scenarios=scenarios)
           
    fig, av = plt.subplots(figsize=(12, 4), layout = 'constrained')
    
    for key in files_dict.keys():
        df_log = pd.read_csv(f'{key}{files_dict[key]}')
        initials = files_dict[key][:2] if files_dict[key][:2].isalpha() else files_dict[key][0]
        if string in list(df_log)[:-1]:
            plt.plot(df_log[string], label = f'{string}_{initials}')
        else:
            raise KeyError(f'You can only plot: {str(list(df_log))[1:-14]}. Not {string}')

    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', borderaxespad=0)
    av.set_ylabel('Time Taken (Seconds)')
    
    xticks = [0]
    for i in range(len(set(df_log['DAY']))):
        xticks.append(48 * (i + 1) - df_log['HH'].loc[0])
  
    labels = []
    for label in list(set(df_log['DAY'])):
        labels.append(f'Day{str(label)}')

    labels.append(f'Day{str(df_log["DAY"].max()+1)}')
    
    av.set_title(f'{PSARunID} | Performance Logs')        
    plt.xticks(xticks, labels, rotation = 'vertical')
    plt.grid(axis = 'x')

    si = ''
    for scenario in scenarios:
        si += f'{get_initials(scenario)}_'
    
    if save:
        return save_result(av, f'{si}Performance_Logs-{string}', directory)
    else:
        return av
'''
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This Function allows you to plot both the % Loading Graph and the Flex_rq graph on the same image, this has sort-of in an unexpected way been replaced by the plot-inter-run flex_reqts'''

def plot_asset_multiple(asset, asset_type, PSARunID, directory=None, scenario='BASE', decorate=True, save=False):
    # -----------------------------------------------------------
    '''Initialise'''
    directory = default_directory(PSARunID, directory=directory)
    config_dict = read_rtp(PSARunID, directory)

    fig, ax = plt.subplots(2, 1, figsize=(10, 5), sharex=True)

    pivot_df = read_flex_rq(asset, PSARunID, directory=directory, scenario=scenario, plot=False, save=False)
    df_line = read_asset(asset, asset_type, PSARunID, directory=directory, scenario=scenario, plot=False, decorate=False, save=False)
    # -----------------------------------------------------------
    '''Plot the required kW graph'''
    try:
        day_max = int(config_dict['DAYS'])
    except KeyError as err:
        day_max = 11

    missing_columns = [number for number in range(0,day_max) if number not in list(pivot_df)]
    for column in missing_columns:
        pivot_df[column] = 0
    pivot_df = pivot_df.loc[:,list(range(0,day_max))]
        
    column_names = convert_to_dates(PSARunID, pivot_df)
    pivot_df.rename(columns = column_names, inplace = True)

    df_melt = pd.melt(pivot_df, var_name='Day', value_name='Req_kw')
        
    date = PSARunID.split('_')[1]
    dt_object = datetime.strptime(date, '%Y-%m-%d-%H-%M')

    for i, date in enumerate(list(pivot_df)):
        auction = auction_relevance(i, dt_object)
        df_date = df_melt[df_melt['Day'] == date]
        if auction == 'Day ahead':
            ax[1].plot(df_date['Req_kw'], label = f'Day-ahead | {date}', linestyle='--', color='Blue')
        elif auction == 'Week ahead':
            ax[1].plot(df_date['Req_kw'], label = f'Week-ahead | {date}', linestyle='-', color='Blue')
        else:
            ax[1].plot(df_date['Req_kw'], label = f'{auction} | {date}', linestyle=':', color='Blue')
        
    xticks = [0]

    for i in range(len(list(pivot_df))):
        xticks.append(48 * (i + 1))
        if decorate:
            ax[1].fill_between(x=(30+(48*i),38+(48*i)), y1=max(pivot_df.max())*1.1, y2=min(pivot_df.min())*1.1, facecolor='green', alpha=0.2)
            ax[1].fill_between(x=(20+(48*i),28+(48*i)), y1=max(pivot_df.max())*1.1, y2=min(pivot_df.min())*1.1, facecolor='yellow', alpha=0.2)

    labels = list(column_names.values())
    end_date = list(column_names.values())[-1].split(' ')[1]
    dt_object = datetime.strptime(end_date, '%d/%m/%y') + timedelta(days = 1)
    labels.append(dt_object.strftime('%A')[:3] + ' ' + dt_object.strftime('%d/%m/%y'))

    ax[1].set_xticks(xticks, labels, rotation = 'vertical')
    ax[1].grid(axis = 'x')
    ax[1].set_ylabel('Required kW')
    # -----------------------------------------------------------
    '''Plot the loading % graph'''
    vmin = get_asset_threshold(asset, PSARunID, directory=directory)

    column_names['Day10.75'] = 'Day10.75'
    df_line.rename(columns = column_names, inplace = True)
        
    df_melt = pd.melt(df_line, var_name='Day', value_name='% Loading')
    df_melt['Day10.75'] = vmin

    date = PSARunID.split('_')[1]
    dt_object = datetime.strptime(date, '%Y-%m-%d-%H-%M')        

    for i, date in enumerate(list(df_line)):
        auction = auction_relevance(i, dt_object)
        df_date = df_melt[df_melt['Day'] == date]
        if auction == 'Day ahead':
            ax[0].plot(df_date['% Loading'], label = f'Day-ahead | {date}', linestyle='--', color='Blue')
        elif auction == 'Week ahead':
            ax[0].plot(df_date['% Loading'], label = f'Week-ahead | {date}', linestyle='-', color='Blue')
        else:
            ax[0].plot(df_date['% Loading'], label = f'{auction} | {date}', linestyle=':', color='Blue')
        
    ax[0].plot(df_melt['Day10.75'],'-', label = f'Threshold: {vmin}%',color='Red')
    xticks = [0]

    for i in range(len(list(df_line))):
        xticks.append(48 * (i + 1))
        if decorate:
            ax[0].fill_between(x=(30+(48*i),38+(48*i)), y1=max(df_line.max())*1.1, facecolor='green', alpha=0.2)
            ax[0].fill_between(x=(20+(48*i),28+(48*i)), y1=max(df_line.max())*1.1, facecolor='yellow', alpha=0.2)

    labels = list(column_names.values())[:-1]
    end_date = list(column_names.values())[-2].split(' ')[1]
    dt_object = datetime.strptime(end_date, '%d/%m/%y') + timedelta(days = 1)
    labels.append(dt_object.strftime('%A')[:3] + ' ' + dt_object.strftime('%d/%m/%y'))

    ax[0].set_ylabel('% Loading')
    ax[0].set_xticks(xticks, labels, rotation = 'vertical')
    ax[0].grid(axis = 'x')
    # -----------------------------------------------------------
    '''Additional Elements for both graphs'''
    legend_elements = [Line2D([0], [0], color='Blue', linestyle='--', label='Day Ahead Auction'),
                        Line2D([0], [0], color='Blue', linestyle='-', label='Week Ahead Auction'),
                        Line2D([0], [0], color='Blue', linestyle=':', label='No Auction Relevance'),
                        Line2D([0], [0], color='Red', linestyle='-', label=f'{vmin}% Threshold')]
                        
    if decorate:
        legend_elements.append(Patch(facecolor='green', alpha=0.2, label='SPM Window'))
        legend_elements.append(Patch(facecolor='yellow', alpha=0.2, label='SEPM Window'))

    ax[0].legend(bbox_to_anchor=(0., 1.02, 1., .102), handles=legend_elements, loc='lower right', ncol=1, borderaxespad=0.)
    ax[0].set_title(f'{PSARunID} | {asset}')
    # -----------------------------------------------------------
    '''Final Steps'''
    si = get_initials(scenario)

    if save:
        return save_result(ax, f'{si}_{asset}_Combination_Graph', director=directory)
    else:
        return ax

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function produces descriptive statistics in a format that is easy to add to a matplotlib graph'''
# How to use this function:
# keys, values = describe_helper(...etc)
# plt.figtext(0.95, 0.49, keys, {'multialignment':'left'})
# plt.figtext(1.05, 0.49, values, {'multialignment':'left'})

def describe_helper(df, column, target_stats=None):
    broken_list = []
    
    df_describe = df[column].describe().loc[target_stats]
    df_describe = df_describe.round(decimals=3)

    possible_targets = str(df_describe).split()[::2]
    if target_stats == None:
        target_stats = possible_targets
    else:
        broken_list = [target for target in target_stats if target not in possible_targets]
     
    if broken_list != []:
        raise ValueError(f'You can only add {(", ").join(possible_targets)} not {(", ").join(broken_list)}')

    splits = str(df_describe).split()[:-4]

    keys, values = '', ''
    for i in range(0, len(splits), 2):
        keys += '{:8}\n'.format(splits[i])
        values += '{:>8}\n'.format(splits[i+1])
    return keys, values


# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function is the same as plot_pl but allows you to enter multiple strings and output one image file with the various plots on it'''

def plot_multiple_pl(strings, PSARunID, directory=None, scenarios=None, save=False):
    directory = default_directory(PSARunID, directory=directory)
    
    if type(strings) != list:
        strings = ['_'.join(string.split(' ')).upper()]
    else:
        strings = ['_'.join(string.split(' ')).upper() for string in strings]
        
    files_dict = find_logs(PSARunID, directory=directory, scenarios=scenarios)
    
    fig, ax = plt.subplots(len(strings), 1, sharex=True)
    
    for q, string in enumerate(strings):
        for key in files_dict.keys():
            df_log = pd.read_csv(os.path.join(key, files_dict[key]))
            initials = files_dict[key][:2] if files_dict[key][:2].isalpha() else files_dict[key][0]
            if string in list(df_log)[:-1]:
                ax[q].plot(df_log[string], label = f'{string}_{initials}')
                
            else:
                raise KeyError(f'You can only plot: {str(list(df_log))[1:-14]}. Not {string}')

        if type(scenarios) != list and scenarios != None and len(strings) == 3:
            # Because of positioning on the graph, this alignment is set up to work only if three strings are selected
            # Selecting more or less will cause things to be out of alignment and possibly un-readable
            keys, values = describe_helper(df_log, string, target_stats=['min','max','mean','std'])
            keys = keys.replace('mean','avg')
            plt.figtext(0.94, 0.64-(q*0.815/len(strings)), keys, {'multialignment':'left'})
            plt.figtext(1.04, 0.64-(q*0.815/len(strings)), values, {'multialignment':'right'})
        
        ax[q].legend(bbox_to_anchor=(1.05, 1), loc='upper left', borderaxespad=0)
        
        xticks = [0]
        for i in range(len(set(df_log['DAY']))):
            xticks.append(48 * (i + 1) - df_log['HH'].loc[0])

        labels = []
        for label in list(set(df_log['DAY'])):
            labels.append(f'Day{str(label)}')
            
        labels.append(f'Day{str(df_log["DAY"].max()+1)}')

        ax[q].set_xticks(xticks, labels, rotation = 'vertical')
        ax[q].grid(axis = 'x')
    
    ax[0].set_title(f'{PSARunID} | Performance Logs')   

    v = len(strings) // 2
    ax[v].set_ylabel('Time Taken (Seconds)')

    si = ''
    if type(scenarios) == list:
        for scenario in scenarios:
            si += f'{get_initials(scenario)}_'

    if save:
        return save_result(ax, f'{si}Performance_Logs_Multiple', directory)
    else:
        return ax

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''almost identical to plot_pl, this overlays each day on top of each other and allows you to ommit certain days. Only accepts one scenario at a time'''
# This function has been replaced by plot_multiple_pl
'''
def plot_pl_daily(string, PSARunID, scenario='BASE', day=None, directory=None, save=False):
    directory = default_directory(PSARunID, directory=directory)
    
    string = '_'.join(string.split(' ')).upper()
    files_dict = find_logs(PSARunID, directory=directory, scenarios=scenario)
    
    for key in files_dict.keys():
        df_log = pd.read_csv(os.path.join(key, files_dict[key]))
        
    df_split = split_days(df_log, string)
    
    if day == None: 
        days = list(df_split)
    elif type(day) == list: 
        days = [f'Day{str(i)}' for i in day]
    else: 
        days = [f'Day{str(day)}']
    
    fig, av = plt.subplots(figsize=(8, 4), layout = 'constrained')
    
    for day in days:
        plt.plot(df_split[day], label = f'{day}')

    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', borderaxespad=0)
    av.set_title(f'{PSARunID} | Performance Logs')
    av.set_ylabel('Time Taken (Seconds)')
    av.set_xlabel('HH No.')
    av.set_xticks([0, 6, 12, 18, 24, 30, 36, 42, 47])

    si = get_initials(scenario)

    if save:
        return save_result(av, f'{si}_Performance_Logs_{string}_DAILY', directory)
    else:
        return av
'''
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Locates the sia data for a particular sheet, day, and HH for a certain PSARunID'''

def query_sia(PSARunID, day, HH, sheet_names = None, directory = None, feeder = None):
    # Get the right folder location
    directory = default_directory(PSARunID, directory)
    
    for item in os.listdir(directory):
        if item[0].isnumeric() and item.split(' - ')[1] == 'SIA_DATA':
            sia = os.path.join(directory, item)
    
    if feeder == None:
        files = os.listdir(sia)
    else:
        for file in os.listdir(sia):
            if '_'.join(file.split('_')[-2:]).removesuffix('.xlsx') == feeder:
                files = [file]
    
    if sheet_names == None:
        sheet_names = ['GEN_MVA','GEN_MW','NET_DEM_MVA','NET_DEM_MW','UND_DEM_MVA','UND_DEM_MW']
    elif type(sheet_names) == list: 
        sheet_names = ['_'.join(sheet_name.split(' ')).upper() for sheet_name in sheet_names]
    else: 
        sheet_names = ['_'.join(sheet_names.split(' ')).upper()]

    dictionary = {}
    
    # Because there's an extra column for time, we need to increase day by 1 and decrease HH due to zero-index
    day_index = day + 1
    HH_index = HH
    
    # For each file in the SIA folder, find the bit of the name that matters...
    for file in files:
        if file.split('_')[1] == 'Feeder':
            row = '_'.join(file.split('_')[-2:]).removesuffix('.xlsx')
        else:
            row = file.split('_')[2]
        
        #... Re-define numbers as nothing to reset each file
        numbers = []
        
        # For each sheet, find the value at the index provided and append it to a list
        for sheet in sheet_names:
            df_test = pd.read_excel(os.path.join(sia, file), sheet_name = sheet)
            numbers.append(round(df_test.iloc[HH_index][day_index], 3))

        # Before moving onto the next file, add it to a dictionary
        dictionary[row] = numbers

    # Convert the resultant dictionary to a dataframe and format
    df = pd.DataFrame.from_dict(dictionary, orient = 'index', columns = sheet_names)
    #df.style.set_caption(f'SIA_Data for: Day {str(day)}, HH {str(HH)}')
    return df

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Plots a graph'''
def sum_primary(primary, PSARunID, directory=None, sheet='NET_DEM_MVA'):
    sia = f'{directory}{const.strSIAfolder}'[:-1]

    if not os.path.isdir(sia):
        print(sia)
        raise KeyError (f'{PSARunID} does not contain SIA_DATA')
    
    if type(primary) != list:
        primaries = [primary]
    else:
        primaries = primary
    
    for primary in primaries:
        files = [file for file in os.listdir(sia) if file.split('_')[1] == 'Feeder' and file.split('_')[2] == primary[:4]]

    if not bool(files):
        raise FileNotFoundError(f'No files could be found in {os.listdir(sia)} matching search criteria {primary[:4]}')
    
    for file in files:
        df1 = pd.read_excel(f'{sia}\\{file}', sheet_name=sheet, index_col=0)
        #print(file)
        #display(df1)
        try:
            df2 = df1.add(df2, fill_value=0)
        except:
            df2 = df1
            
    return df2

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Takes the sum of a primary and divides it by the rating as per sia, using a BSP_file like:  bsps_primary_txs_operating_Oxfordshire_trial_Maximo_SIA_RATINGS_v9.xlsx'''
def identify_threshold(primary, PSARunID, BSP_file, directory=None, sheet='NET_DEM_MVA'):
    max_loading_actual, max_loading_potential = 0, 0
    
    df1 = sum_primary(primary, PSARunID, directory=directory, sheet=sheet)
    df2 = pd.read_excel(BSP_file, usecols=[1,2])
    df2['Transformer_name'] = df2['Transformer_name'].apply(lambda x: x[:4] if len(x) == 9 else x)
    #df2.replace(np.nan, 7, inplace=True)
    # Usecols = 'BSP','Transformer_name','Rating to be used by SIA in dashboard \n(alerts)'
    df3 = df2.groupby(['Transformer_name']).sum()

    # df3 here includes the transformer code as transformer_name e.g BERI, COLO, KENN, etc.
    # df3 also includes Cowley Local Reserve and Cowley local main as COLO_A1MTA, COLO_A1MTB, COLO_A2MTA, COLO_A2MTB but it doesn't do anything with them
    #display(df1)
    
    if type(primary) != list:
        primaries = [primary]
    else:
        primaries = primary
        
    for primary in primaries:
        if df3.at[primary, 'Rating_MVA'] > 0:
            max_loading_potential += df3.at[primary, 'Rating_MVA']
        else:
            raise ValueError(f'Substation {primary} has rating of 0')
    
    df_threshold = df1/max_loading_potential * 100
    return df_threshold

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Plots a graph of the sia, summed by primary, as % loading'''
def plot_sia(primary, PSARunID, directory=None, sheet='UND_DEM_MVA', threshold=0, decorate=True, plot=True, save=False, BSP_file=None):
    directory = default_directory(PSARunID, directory)
    if not BSP_file:
        config_dict = ut.PSA_SND_read_config(const.strConfigFile)
        BSP_file = f'{config_dict[const.strWorkingFolder]}\\packages\\psa\\bsps_primary_txs_RATINGS_Thresholds_v1.xlsx'

    fixed_data = {'COLO_BSP_MAIN-Rating':180, # All of these are Fixed values in MVA
                  'COLO_BSP_RESERVE-Rating':180, 
                  'COLO-Demand':24.75, 
                  'UNIS-Demand':6.33, 
                  'PRSC-Demand':18.53}
    
    if primary == 'COLO_A1MTA' or primary == 'COLO_A2MTA':
        df_sia = sum_primary(['BERI','KENN','WALL'], PSARunID, directory, sheet)
        df_sia = df_sia / fixed_data['COLO_BSP_RESERVE-Rating'] * 100
        title = f'SIA Data | COLO_BSP_RESERVE'
        colour = 'darkred'
    elif primary == 'COLO_A1MTB' or primary == 'COLO_A2MTB':
        df_sia = sum_primary('ROSH', PSARunID, directory, sheet)
        df_sia += fixed_data['COLO-Demand']
        df_sia += fixed_data['UNIS-Demand']
        df_sia += fixed_data['PRSC-Demand'] 
        df_sia = df_sia / fixed_data['COLO_BSP_MAIN-Rating'] * 100
        title = f'SIA Data | COLO_BSP_MAIN'
        colour = 'darkred'
    elif primary[:4] in ['BERI','ROSH','KENN','WALL']:
        df_sia = identify_threshold(primary[:4], PSARunID, BSP_file, directory=directory, sheet=sheet)
        title = f'SIA Data | {primary[:4]} EXL5'
        colour = 'Navy'
    elif primary[-1].isnumeric():
        df_sia = pd.read_excel(f'{directory}{const.strBfolder}LN_LDG_DATA\\PSA_B_Loading-{primary}.xlsx', index_col=0)
        title = f'LN Data | {primary}'
        colour = 'purple'
    else:
        df_sia = pd.read_excel(f'{directory}{const.strBfolder}TX_LDG_DATA\\PSA_B_Loading-{primary}.xlsx', index_col=0)
        title = f'TX Data | {primary}'
        colour = '#E69646'

    column_names = convert_to_dates(PSARunID, df_sia)

    if type(threshold) == float or type(threshold) == int and threshold > 0:
        vmin = threshold
        column_names['Day10.75'] = 'Day10.75'
        df_sia.rename(columns = column_names, inplace = True)
        df_melt = pd.melt(df_sia, var_name='Day', value_name='pct_loading')
        df_melt['Day10.75'] = vmin
        threshold = True
    elif type(threshold) == str:
        df_describe = df_sia.describe()
        try:
            if threshold == 'min':
                vmin = round(df_describe.loc[threshold].min(),3)
            elif threshold == 'max':
                vmin = round(df_describe.loc[threshold].max(),3)
            else:
                vmin = round(df_describe.loc[threshold].mean(),3)
        except:
            raise KeyError(f'String Threshold can only be {(", ").join(df_describe.index[1:])} not {threshold}')
        column_names['Day10.75'] = 'Day10.75'
        df_sia.rename(columns = column_names, inplace = True)
        df_melt = pd.melt(df_sia, var_name='Day', value_name='pct_loading')
        df_melt['Day10.75'] = vmin
        threshold = True      
    else:
        df_sia.rename(columns = column_names, inplace = True)
        df_melt = pd.melt(df_sia, var_name='Day', value_name='pct_loading')
        threshold = False

    old_threshold = get_asset_threshold(primary, PSARunID, directory=directory)
    df_melt['Day11.75'] = old_threshold
    
    date = PSARunID.split('_')[1]
    dt_object = datetime.strptime(date, '%Y-%m-%d-%H-%M')
    
    if plot:
        fig, av = plt.subplots(figsize=(12, 4)) 

        for i, date in enumerate(list(df_sia)):
            auction = auction_relevance(i, dt_object)
            df_date = df_melt[df_melt['Day'] == date]
            if auction == 'Day ahead':
                plt.plot(df_date['pct_loading'], label = f'Day-ahead | {date}', linestyle='--', color=colour)
            elif auction == 'Week ahead':
                plt.plot(df_date['pct_loading'], label = f'Week-ahead | {date}', linestyle='-', color=colour)
            else:
                plt.plot(df_date['pct_loading'], label = f'{auction} | {date}', linestyle=':', color=colour)

        if threshold:
            plt.plot(df_melt['Day10.75'],'-', label = f'Threshold: {vmin}%',color='Red')
            labels = list(column_names.values())[:-1]
            plt.plot(df_melt['Day11.75'],':', label = f'Current Threshold: {old_threshold}%', color='Red', alpha=0.5)
        else:
            labels = list(column_names.values())
    
        xticks = [0]

        for i in range(len(list(df_sia))):
            xticks.append(48 * (i + 1))
            if decorate:
                y1 = max([max(df_sia.max()), vmin, old_threshold]) if threshold else max(df_sia.max())
                plt.fill_between(x=(30+(48*i),38+(48*i)), y1=y1*1.1, facecolor='green', alpha=0.2) # SPM Window
                plt.fill_between(x=(20+(48*i),28+(48*i)), y1=y1*1.1, facecolor='yellow', alpha=0.2) # SEPM Window

        end_date = list(column_names.values())[-2].split(' ')[1]
        dt_object = datetime.strptime(end_date, '%d/%m/%y') + timedelta(days = 1)
        labels.append(dt_object.strftime('%A')[:3] + ' ' + dt_object.strftime('%d/%m/%y'))

        keys, values = describe_helper(df_melt, 'pct_loading', target_stats=['min','max','mean','std'])
        keys = keys.replace('mean','avg')
        plt.figtext(0.14, 0.86, keys, {'multialignment':'left'})
        plt.figtext(0.24, 0.86, values, {'multialignment':'right'})
                
        legend_elements = [Line2D([0], [0], color=colour, linestyle='--', label='Day Ahead Auction'),
                        Line2D([0], [0], color=colour, linestyle='-', label='Week Ahead Auction'),
                        Line2D([0], [0], color=colour, linestyle=':', label='No Auction Relevance'),
                        Line2D([0], [0], color='Red', alpha=0.5, linestyle=':', label=f'{old_threshold}% Old Threshold')]
        
        if threshold:
            legend_elements.append(Line2D([0], [0], color='Red', linestyle='-', label=f'{vmin}% Threshold'))
        if decorate:
            legend_elements.append(Patch(facecolor='green', alpha=0.2, label='SPM Window'))
            legend_elements.append(Patch(facecolor='yellow', alpha=0.2, label='SEPM Window'))

        av.legend(bbox_to_anchor=(0., 1.02, 1., .102), handles=legend_elements, loc='lower right', ncol=1, borderaxespad=0.)
        av.set_ylabel('% Loading')
        
        if threshold:
            constraints = df_melt[df_melt['pct_loading'] > vmin].count(numeric_only=True).iloc[0]
            av.set_title(f'{title} | {constraints} Import Constraints')
        else:
            av.set_title(title)

        plt.xticks(xticks, labels, rotation = 'vertical')
        plt.grid(axis = 'x')
        if save:
            return save_result(av, f'{primary}_SIA_LOADING_PCT_GRAPH', directory), vmin
        else:
            return av
    
    else:
        df_melt = pd.melt(df_sia, var_name='Day', value_name='pct_loading')
        if save:
            return save_result(df_melt, f'{primary}_SIA_LOADING_PCT_DATA', directory), vmin
        else:
            return df_melt

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function tests to ensure that the nominal_current_adj part of the constraints.xlsx file varies according to differing loading_thresholds'''
# This function is incompatable with latest versions
'''
def constraints_validation(PSARunID, directory=None):
    constraints = pd.read_excel(os.path.join(directory, f'{PSARunID}{const.strPSAConstraints}.xlsx'), usecols=[11,12,13])
    # Use Cols: loading_thershold_pct, nominal_current_kA, nominal_current_adj_kA
    constraints['test'] = constraints['loading_threshold_pct'] * constraints['nominal_current_kA'] / 100
    constraints['validation'] = round(constraints['test'],10) == round(constraints['nominal_current_adj_kA'],10)
    if constraints[constraints['validation'] == False].shape[0] == 0:
        print('Constraints validated successfully, adjusted current is accurate to the loading threshold')
        return None
    else:
        print('Issues found')
        return constraints[constraints['validation'] == False]
'''
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function takes all columns of a dataframe and renames them with a tag on the end, useful when merging dataframes to know where data has come from'''
def add_tag(df, tag):
    names_0 = list(df)
    names_1 = [f'{column}-{tag}' for column in names_0]
    new_dict = {names_0: names_1 for names_0, names_1 in zip(names_0, names_1)}
    df.rename(columns=new_dict, inplace=True)
    return df

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function checks to see if there are any duplicate column names, according to a regular expression, in a dataframe'''

def validation(df, regex=None):
    if regex == None:
        uniques = []
        duplicates = []
        column_names = [column[:-4] for column in list(df)]
        for column in column_names:
            if column in uniques:
                duplicates.append(column)
            else:
                uniques.append(column)
    else:
        duplicates = [regex]
    
    for regex in duplicates:
        columns = [column for column in list(df) if regex in column]
        i = 1
        while i < len(columns):
            if list(df[columns[0]]) == list(df[columns[i]]):
                df.drop(columns[i], axis=1, inplace=True)
            else:
                print(f'"{columns[0]}" and "{columns[i]}" failed to validate correctly')
            i += 1
    return df

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function takes all the relevant data from constraints, flex_reqts, sf_responses, and snd_responses and then merges them into one big dataframe, highlighting any discrepencies if validate=true'''

def read_sensitivity(PSARunID, directory=None, validate=True):
    directory = default_directory(PSARunID, directory)
    file_names = [f'{const.strPSAConstraints}', f'{const.strPSAFlexReqts}', f'{const.strPSAResponses}-V0', f'{const.strSNDResponses}-V0']

    constraints = pd.read_excel(os.path.join(directory, f'{PSARunID}{file_names[0]}.xlsx'), index_col=0, usecols=[0,10,11,14])
    # Use cols: req_id, loading_pct, loading_thershold_pct, scenario
    constraints = add_tag(constraints, 'con')
    flex_reqts = pd.read_excel(os.path.join(directory, f'{PSARunID}{file_names[1]}.xlsx'), index_col=0, usecols=[0,8,15])
    # Use cols: req_id, loading_threshold_pct, required_power_kw
    flex_reqts = add_tag(flex_reqts, 'flx')
    sf_responses = pd.read_excel(os.path.join(directory, f'{PSARunID}{file_names[2]}.xlsx'), index_col=0, usecols=[0,1,2,3,4,5,6,7,12,13,14,15,16,21,22])
    # Use cols: req_id, resp_id, bsp, primary, feeder, secondary, terminal, constrained_pf_id, start_time, duration, required_power_kw, flex_pf_id, offered_power_kw, sensitivity_factor, sensitivity_factor_dir
    sf_responses = add_tag(sf_responses, 'sfs')
    snd_responses = pd.read_excel(os.path.join(directory, f'{PSARunID}{file_names[3]}.xlsx'), index_col=0, usecols=[0,14,16])
    # Use cols: req_id, required_power_kW, offered_power_kW
    snd_responses = add_tag(snd_responses, 'snd')

    super_df = pd.concat([sf_responses, constraints, flex_reqts, snd_responses], axis=1, join='inner')
    
    if validate == True:
        super_df = validation(super_df)
    
    return super_df

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function uses the PSARunID to work out which day of the week each day0, day1, day,2 etc is talking about and if they are actually relevant for auctions'''

def auction_relevance(dt_object, start_date):
    if type(dt_object) == int:
        dt_object = start_date + timedelta(days=dt_object)
    elif type(dt_object) == str:
        dt_object = start_date + timedelta(days=int(dt_object[3:]))
    # Format should either be str 'day0' 'day1'... etc or int 0, 1, 2... etc    
    
    if dt_object.weekday() >= 5:
        relevance = 'Weekend'
    elif start_date + timedelta(days=13-start_date.weekday()) >= dt_object >= start_date + timedelta(days=7-start_date.weekday()):
        relevance = 'Week ahead'
    elif dt_object.weekday() == start_date.weekday() + 1: 
        relevance = 'Day ahead'
    else:
        relevance = 'None'

    return relevance

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function calculates the difference between two date_times in days.'''

def delta(date1, start_date):
    return int((date1.date()-start_date.date()).days)

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function takes a dataframe and returns the values for one particular group in the order they appear in the index.'''

def groupby_values(df, on='sensitivity_factor', how='mean'):
    asset_dict = {}
    # pf_id = [] for debugging
    values = []

    unique_assets = list(set(df['constrained_pf_id']))
    unique_flex = list(set(df['flex_pf_id']))

    how = how.upper()
    
    for asset in unique_assets:
        for flex in unique_flex:
            asset_dict[flex] = np.nan

        asset_df = df[df['constrained_pf_id'] == asset]
        if how=='MEAN':
            filtered_df = asset_df[['flex_pf_id',on]].groupby(['flex_pf_id']).mean()
        elif how=='SUM':
            filtered_df = asset_df[['flex_pf_id',on]].groupby(['flex_pf_id']).sum()
        elif how=='COUNT':
            filtered_df = asset_df[['flex_pf_id',on]].groupby(['flex_pf_id']).count()
        else:
            raise ValueError ('how must be either mean, sum, or count')

        for index in list(filtered_df.index):
            asset_dict[index] = round(filtered_df.loc[index][0],6)

        #pf_id = pf_id + list(flex_dict.keys()) for debugging
        values = values + list(asset_dict.values())
    
    return values

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function uses the groupby function to create a multiindexed dataframe, helping to give a better idea regarding sensitity factors'''
# Updated and untested scenario used to be base, if statement for scenario non-existant

def sf_analysis(PSARunID, directory=None, scenario='BASE', decorate=True, save=False):
    directory = default_directory(PSARunID, directory)
    scenario = '_'.join(scenario.split(' ')).upper()
    sf_dict = get_sf_dict(PSARunID, directory)
    try:
        scenario_to_version = sf_dict[scenario]
    except KeyError:
        raise FileNotFoundError(f'Scenario {scenario} Not Found in any SF batches. Scenarios present: {list(sf_dict.items())}')
    df = []
    for version in scenario_to_version:
        df_temp = pd.read_excel(f'{directory}\\{PSARunID}{const.strPSAResponses}-V{version}.xlsx', usecols=['req_id','resp_id','constrained_pf_id','scenario',\
                                                                                                            'start_time','duration','required_power_kw','flex_pf_id',\
                                                                                                            'offered_power_kw','sensitivity_factor','sensitivity_factor_dir'])
        try:
            df = pd.concat([df, df_temp])
        except:
            df = df_temp.copy()

    if len(df['scenario'].drop_duplicates()) != 1:
        df = df[df['scenario'] == scenario]
    
    # This could be streamlined as uniques* are defined in the groupby function as well
    unique_assets = list(set(df['constrained_pf_id']))
    unique_flex = list(set(df['flex_pf_id']))

    # To flip the groupby, swap the order of iterables and then the names of the index in the next two lines
    iterables = [unique_assets,unique_flex]
    index = pd.MultiIndex.from_product(iterables, names=('constrained_pf_id','flex_pf_id'))

    values1 = groupby_values(df, on='offered_power_kw', how='sum')
    values2 = groupby_values(df, on='sensitivity_factor', how='mean')
    values3 = groupby_values(df, on='offered_power_kw', how='count')
    values4 = groupby_values(df, on='sensitivity_factor_dir', how='mean')
    # Groupby count no of responses/HHs provided for

    stacked_df = pd.DataFrame({'Scenario': scenario,'HHs w/ offer': values3, 'Sum Offered kW': values1, 'Avg SF': values2, 'Avg SF dir': values4}, index=index)

    if decorate:
        stacked_df = stacked_df.style.background_gradient(subset='Avg SF', cmap='coolwarm_r', low=0, text_color_threshold=0.35, vmin=0, vmax=1)\
                    .highlight_between(subset='Avg SF', left=-999, right=0, color='black', inclusive='left')\
                    .highlight_null(subset='Avg SF', null_color='white')\
                    .format(formatter= {'HHs w/ offer': '{:.0f}', 'Sum Offered kW': '{:.0f}', 'Avg SF dir': '{:.0f}'})\
                    .set_properties(**{'text-align': 'right'})

    si = get_initials(scenario)

    if save:
        return save_result(stacked_df, f'{si}_SF_Overview', directory=directory)
    else:
        return stacked_df

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This Function returns the batch and version number of a file either as a string or two integer values'''
def get_sf_dict(PSARunID, directory):
    directory = default_directory(PSARunID, directory)
    sf_files = [file for file in os.listdir(directory) if const.strPSAResponses in file]
    if not sf_files:
        raise FileNotFoundError(f'No SF Files in {directory}\\{PSARunID}')
        
    scenario_batch_dict = {}
    for file in sf_files:
        batch = file[-6]
        scenario = pd.read_excel(f'{directory}\\{file}', usecols=['scenario'])['scenario'].drop_duplicates().to_list()
        #if len(scenario) > 1:
        #    raise TypeError(f'get_sf_dict() generates an invalid response: {scenario}')
        for s in scenario:
            if s not in scenario_batch_dict.keys():
                scenario_batch_dict[s] = [batch]
            else:
                scenario_batch_dict[s].append(batch)

    return scenario_batch_dict

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function returns the overview of all the SF_responses files and what information lies in which file'''
def read_sf_by_batch(PSARunID, directory, save=False):
    directory = default_directory(PSARunID, directory)
    sf_files = [file for file in os.listdir(directory) if const.strPSAResponses in file]

    if not sf_files:
        raise FileNotFoundError(f'No SF Files in {directory}\\{PSARunID}')

    df_old = []
    for file in sf_files:
        df_new = pd.read_excel(f'{directory}\\{file}', usecols=['req_id','resp_id','constrained_pf_id','scenario','required_power_kw','offered_power_kw','sensitivity_factor'])
        df_new['batch'] = int(file.split('V')[1][0])
        try:
            df_old = pd.concat([df_new, df_old], ignore_index=True)
        except:
            df_old = df_new.copy()

    df_old = df_old.groupby(['batch','req_id','resp_id'])[['scenario','constrained_pf_id','required_power_kw','offered_power_kw','sensitivity_factor']].max()
    
    if save:
        return save_result(df_old, f'SF_batch_info', directory)
    else:
        return df_old
    
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function converts an int to it's equivelant HH timestamp. 0 -> 00:00 | 47 -> 23:30'''

def HH_to_timestamp(value):
    return f'{str(int(value) // 2)}:{str((int(value) % 2)*3)}0'

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function is very similar to plot_asset but rather than plotting loading % it plots required kW. In a future version plot_asset and read_flex_rq could likely be merged'''

def read_flex_rq(asset, PSARunID, directory=None, scenario='BASE', plot=True, save=False):
    directory = default_directory(PSARunID, directory=directory)
    
    scenario = '_'.join(scenario.split(' ')).upper()
    si = get_initials(scenario)
    
    flex_reqts = find_data(PSARunID, 'FLEX_REQTS', directory=directory, scenario=scenario)
    super_df = pd.read_excel(os.path.join(flex_reqts, f'{PSARunID}{const.strPSAFlexReqts}.xlsx'))

    date = PSARunID.split('_')[1]
    dt_object = datetime.strptime(date, '%Y-%m-%d-%H-%M')

    super_df['start_time'] = pd.to_datetime(super_df['start_time'])
    super_df['day'] = super_df['start_time'].apply(delta, start_date=dt_object)
    super_df['HH'] = super_df['start_time'].dt.hour * 2 + super_df['start_time'].dt.minute / 30
    
    asset_df = super_df[super_df['constrained_pf_id'] == asset]
    pivot_df = asset_df[['required_power_kw','day','HH']].pivot(index='HH',columns='day', values='required_power_kw')
    if pivot_df.shape[0] != 48:
        missing_values = [value for value in list(range(0, 48)) if value not in list(pivot_df.index.to_series())]
        missing_dict = {}

        for value in missing_values:
            missing_dict[value] = [0]*pivot_df.shape[1]

        filler_df = pd.DataFrame.from_dict(missing_dict, orient='index', columns=list(pivot_df))
        pivot_df = pd.concat([pivot_df, filler_df])
        pivot_df.sort_index(inplace=True)
    
    pivot_df.set_index(pivot_df.index.map(HH_to_timestamp), inplace=True)
    pivot_df.fillna(0, inplace=True)
    #display(pivot_df) for debugging
    
    if plot:
        config_dict = read_rtp(PSARunID, directory)
        try:
            day_max = int(config_dict['DAYS'])
        except KeyError as err:
            day_max = 11

        missing_columns = [number for number in range(0,day_max) if number not in list(pivot_df)]
        for column in missing_columns:
            pivot_df[column] = 0
        pivot_df = pivot_df.loc[:,list(range(0,day_max))]

        column_names = convert_to_dates(PSARunID, pivot_df)
        pivot_df.rename(columns = column_names, inplace = True)

        df_melt = pd.melt(pivot_df, var_name='Day', value_name='Req_kw')

        title = f'{PSARunID}: {asset}'
        ax = plot_generic_graph(PSARunID, df_melt, title=title, y_label='Req kW', decorate='type_1', describer=True)
        
        if save:
            return save_result(ax, f'{si}_{asset}_REQUIRED_KW_GRAPH', directory)
        else:
            return ax
    else:
        config_dict = read_rtp(PSARunID, directory)
        try:
            day_max = int(config_dict['DAYS'])
        except KeyError as err:
            day_max = 11

        missing_columns = [number for number in range(0,day_max) if number not in list(pivot_df)]
        for column in missing_columns:
            pivot_df[column] = 0
        pivot_df = pivot_df.loc[:,list(range(0,day_max))]

        column_names = convert_to_dates(PSARunID, pivot_df)
        pivot_df.rename(columns = column_names, inplace = True)

        if save:
            return save_result(pivot_df, f'{si}_{asset}_REQUIRED_KW_DATA', directory)
        else:
            return pivot_df

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function checks the sf_calc intermediary files and makes sure that the initial loading_pct figure matches that of the constraints file for the same req_id. It returns a dataframe.'''

def sf_debug(PSARunID, directory=None, version_number='latest', decorate=True):
    directory = default_directory(PSARunID, directory=directory)
    fldr = f'{directory}\\TEMP_SF_FILES'

    if version_number == 'latest':
        fldr = latest_version(fldr)
    else:
        fldr = f'{fldr}\\V{version_number}'

    sf_files = [x for x in os.listdir(fldr) if x[0:7] == 'SF_calc']

    d = {'req_id':[], 'resp_id':[], 'Loading_pct-debug':[]}
    old_df = pd.DataFrame(data=d)

    for file in sf_files:
        new_df = pd.read_csv(os.path.join(fldr, file), usecols=['Name_network_element','Loading_pct'], nrows=1)
        new_df = add_tag(new_df, 'sfs')
        new_df['req_id'] = int(file.split('_')[-2])
        new_df['resp_id'] = file.split('_')[-1].removesuffix('.csv')
        old_df = pd.concat([old_df, new_df])

    df = pd.read_excel(os.path.join(directory, f'{PSARunID}{const.strPSAConstraints}.xlsx'), usecols=[0,6,7,8,10,11,14], converters={'start_time': pd.to_datetime})
    df.rename({'loading_pct': 'loading_pct-con'}, axis=1, inplace=True)
    # Use cols: req_id, start_time, loading_pct, scenario

    type_scenario_dict = {const.strPFassetTrans:'TX_LDG_DATA\\',
                          const.strPFassetLine:'LN_LDG_DATA\\',
                          'BASE': const.strBfolder,
                          'MAINT': const.strMfolder,
                          'MAINT_CONT': const.strMCfolder,
                          'CONT': const.strCfolder}

    start_time_now = datetime.strptime(PSARunID.split('_')[1], '%Y-%m-%d-%H-%M')
    start_time = datetime(*start_time_now.timetuple()[:3])
    df['start'] = (df['start_time']-start_time).dt.days*48 + [(df.loc[i,'start_time'].timetuple()[3]*2) + (df.loc[i,'start_time'].timetuple()[4]//30) for i in range(len(df))]
    df['day'] = [f'Day{start // 48}' for start in df['start']]
    df['hh'] = [HH_to_timestamp(start % 48) for start in df['start']]

    df['loading_pct-map'] = np.nan
    for asset in df['constrained_pf_id'].unique():
        df_mask = df[df['constrained_pf_id'] == asset].reset_index()
        scenario = type_scenario_dict[df_mask.loc[0,"scenario"]]
        asset_type = type_scenario_dict[df_mask.loc[0,"constrained_pf_type"]]
        initials = get_initials(df_mask.loc[0,"scenario"])
        file = f'{directory}{scenario}{asset_type}PSA_{initials}_Loading-{asset}.xlsx'
        df_temp = pd.read_excel(file, index_col=0)
        for req_id, day, hh in zip(df_mask['req_id'],df_mask['day'],df_mask['hh']):
            df.loc[req_id, 'loading_pct-map'] = df_temp.loc[hh, day]

    merged_df = pd.merge(df, old_df, how='inner', on='req_id')

    a, b, c = [round(merged_df['Loading_pct-sfs'],2).to_list(), round(merged_df['loading_pct-con'],2).to_list(), round(merged_df['loading_pct-map'],2).to_list()]
    merged_df['validation'] = all([a == b, b == c])

    merged_df = merged_df.loc[:,['req_id','resp_id','Name_network_element-sfs','start_time','scenario','loading_threshold_pct','Loading_pct-sfs','loading_pct-con','loading_pct-map','validation']]
    if decorate:
        merged_df = merged_df.style.apply(lambda rows: [f'background: yellow' if not rows.validation else '' for row in rows], axis=1)

    return merged_df

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function plots all assets and there maintanence period This function is depreacted and has been replaced by plot_gantt_outages'''
'''
def plot_outages(PSARunID, directory=None, asset=None, scenario='MAINT', save=False):
    directory = default_directory(PSARunID, directory)
    initials = get_initials(scenario)
    asset_file = ''

    asset_outage_tables = f'{directory}{const.strINfolder}{initials}-ASSET_OUTAGE_TABLES'
    
    data_dict = {}
    assets = []
    i = 1

    if asset == None:
        for folder in os.listdir(asset_outage_tables):
            assets += [file.removesuffix('.xlsx') for file in os.listdir(f'{asset_outage_tables}\\{folder}')]
    elif type(asset) != list:
        assets = [asset]
    else:
        assets = asset

    for folder in os.listdir(asset_outage_tables):
        for file in os.listdir(f'{asset_outage_tables}\\{folder}'): # Read each asset in both lines and transformers
            if file.removesuffix('.xlsx') in assets:
                data = []
                asset_file = f'{asset_outage_tables}\\{folder}\\{file}'
                df = pd.read_excel(asset_file, index_col = 0)
                for column in list(df):
                    data = data + list(df[column])

                data_dict[file.removesuffix('.xlsx')] = [value*i for value in data]
                i += 0.2

    df_joined = pd.DataFrame.from_dict(data_dict)

    df_joined['index'] = np.arange(0,len(df_joined),1)
    df_joined['DAY'] = df_joined['index'] // 48

    try:
        dates_dict = convert_to_dates(PSARunID, df)
    except:
        raise TypeError ('Asset Not Found')

    labels = list(dates_dict.values())

    end_date = list(dates_dict.values())[-1].split(' ')[1]
    dt_object = datetime.strptime(end_date, '%d/%m/%y') + timedelta(days = 1)
    labels.append(dt_object.strftime('%A')[:3] + ' ' + dt_object.strftime('%d/%m/%y'))

    xticks = [0]

    for i in range(len(set(df_joined['DAY']))):
        xticks.append(48 * (i + 1))
        
    final_dict = dict(zip(xticks, labels))
    
    fig, av = plt.subplots(figsize=(12, 4))
    
    i = 1
    if len(assets) == 1:
        df_asset = df_joined[[asset, 'index']]
        df_asset['delta'] = df_asset[asset].diff()
        df_delta = df_asset[df_asset['delta']!=0]
       
        y = df_joined[asset]
        x = df_joined['index']
        start = df_delta.iloc[1].name
        
        av.plot(x, y)
        av.text(x[start], y[start], f'{asset}', fontsize='small')
        
        timings_dict = {}
        for q in range(1, df_delta.shape[0], 1):
            start = (df_delta.iloc[q].name)
            timings_dict[start] = HH_to_timestamp(start % 48)
        
        av2 = av.twiny()
        av2.set_xlim(av.get_xlim())
        av2.set_xticks(list(timings_dict.keys()), list(timings_dict.values()))
        av.set_title(f'{assets[0]} | Outage Data')

    else:
        for asset in assets:
            y = df_joined[asset]
            x = df_joined['index']
            start = df_joined[df_joined[asset] == i].iloc[1].name

            av.barh(x, y)
            av.text(x[start], y[start], f'{asset}', fontsize='small')

            i += 0.2
            
        av.set_title(f'{PSARunID} | Outage Data')
        
    av.set_ylabel('Out of Service')
    #av.xaxis.set_minor_locator(MultipleLocator(12)) # requires from matplotlib.ticker import (MultipleLocator, AutoMinorLocator)
    
    av.set_yticks([0,i*1.2], [])
    av.set_xticks(list(final_dict.keys()), list(final_dict.values()), rotation = 'vertical')
    av.grid(axis = 'x')
    
    if save:
        return save_result(av, f'{PSARunID}-OUTAGE_DATA', directory)
    else:
        return av
'''
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This Function plots outages directly from a maint file  This function is depreacted and has been replaced by plot_gantt_outages_by_file'''
'''
def plot_file_outages(directory, no_of_days=10, save=False, plot=True):
    directory = ('\\').join(directory.split('/')) # For Coherency
    df = pd.read_excel(directory, converters={'START_OUTAGE': pd.to_datetime,
                                              'END_OUTAGE': pd.to_datetime})

    start_time = datetime.now()
    start_time = datetime(*start_time.timetuple()[:3])
    end_time = start_time + timedelta(days=no_of_days)

    values = []
    for i in df.index:
        start_relevance = is_time_between(start_time, end_time, df.loc[i,'START_OUTAGE'])
        end_relevance = is_time_between(start_time, end_time, df.loc[i,'END_OUTAGE'])
        values.append(start_relevance or end_relevance)

    df['RELEVANT'] = values
    filtered_df = df[df['RELEVANT'] == True]

    fake_PSARunID = f'ff-{datetime.strftime(start_time, "%Y-%m-%d-%H-%M")}'
    dates_dict = convert_to_dates(fake_PSARunID, range(0,no_of_days))
    labels = list(dates_dict.values())

    end_date = list(dates_dict.values())[-1].split(' ')[1]
    dt_object = datetime.strptime(end_date, '%d/%m/%y') + timedelta(days = 1)
    labels.append(dt_object.strftime('%A')[:3] + ' ' + dt_object.strftime('%d/%m/%y'))

    xticks = [0]

    for i in range(0,no_of_days):
        xticks.append(48 * (i + 1))

    final_dict = dict(zip(xticks, labels))

    if plot:
        fig, av = plt.subplots(figsize=(12, 4))
        i = 1

        for v in filtered_df.index:
            start_out = filtered_df.loc[v,'START_OUTAGE']
            end_out = filtered_df.loc[v,'END_OUTAGE']
            asset = filtered_df.loc[v,'ASSET_ID']
            data = []
            for hh in range(0,48*no_of_days):
                data.append(i if is_time_between(start_out, end_out, start_time + timedelta(hours=hh/2)) else 0)
            df_asset = pd.DataFrame(data=data, columns=[asset])

            y = df_asset[asset]
            x = list(df_asset.index)

            start = df_asset[df_asset[asset] == i].iloc[0].name 
            av.plot(x, y)
            av.text(x[start], y[start], f'{asset}', fontsize='small')
            i += 0.2

        y, z = 0, 0
        for p, date in enumerate(list(final_dict.values())):
            auction = auction_relevance(p, start_time)
            if auction == 'Day ahead' and y == 0:
                plt.fill_between(x=(p*48,(p+1)*48), y1=i*1.2, y2=0, facecolor='blue', alpha=0.1)
                plt.text(p*48 + 24, i, '- Day -\nAhead', color='blue', alpha=0.25, horizontalalignment='center', fontsize='large')
                y += 1
            elif auction == 'Week ahead':
                z += 1
                if z == 1:
                    plt.fill_between(x=(p*48,(p+1)*48), y1=i*1.2, y2=0, facecolor='blue', alpha=0.1)
                    plt.text(p*48 + 24, i, '- Week -\nAhead', color='blue', alpha=0.25, horizontalalignment='center', fontsize='large')
                elif z <= 5:
                    plt.fill_between(x=(p*48,(p+1)*48), y1=i*1.2, y2=0, facecolor='blue', alpha=0.1)
                    plt.text(p*48 + 24, i*1.04, '--->', color='blue', alpha=0.25, horizontalalignment='center', fontsize='large')


        file_name = directory.split('\\')[-1].removesuffix('.xlsx')
        av.set_title(f'{file_name} | Outage Data')

        av.set_ylabel('Out of Service')

        av.set_yticks([0,i*1.2], [])
        av.set_xticks(list(final_dict.keys()), list(final_dict.values()), rotation = 'vertical')
        av.grid(axis = 'x')
    else:
        return filtered_df, final_dict

    if save:
        save_directory = ('\\').join(directory.split('\\')[:-1])
        return save_result(av, f'{file_name}_{no_of_days}_day_plot', save_directory)
    else:
        return av
'''
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Function used by plot_gantt outages to determine colour of lines, trafos, and synms'''
def colour(row):
    c_dict0 = {const.strAssetLine:'#E64646',const.strAssetTrans:'#E69646',const.strAssetGen:'#34D05C',const.strAssetLoad:'#f51197'}
    c_dict1 = {'MAINT':'#E69646','CONT':'#E64646'}
    c_dict2 = {'BASE':'#34D05C','MAINT':'#E69646','CONT':'#E64646','MAINT_CONT':'#f51197'}
    c_dict3 = {'BASE-OPEN':'#6aa84f','BASE-CLOSED':'#d9ead3',
               'MAINT-OPEN':'#e69138','MAINT-CLOSED':'#fce5cd',
               'CONT-OPEN':'#cc0000','CONT-CLOSED':'#f4cccc',
               'MAINT_CONT-OPEN':'#a64d79','MAINT_CONT-CLOSED':'#ead1dc'}
    
    try:
        columns = row.index.to_list()
        if 'SCENARIO' in columns and 'ACTION' in columns:
            colour = c_dict3[f"{row['SCENARIO']}-{row['ACTION']}"]
        elif 'SCENARIO' in columns:
            colour = c_dict2[row['SCENARIO']]
        elif 'FILE' in columns:
            colour = c_dict1[row['FILE']]
        elif 'ASSET_TYPE' in columns:
            colour = c_dict0[row['ASSET_TYPE']]
        else:
            colour = 'silver'
    except KeyError as err:
        print(err)
        colour = 'silver'
    return colour

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Very similar to plot_file_outages but displays result as a gantt chart rather than a line graph'''
def plot_gantt_outages(PSARunID, directory, no_of_days=11, start_time='now', save=False, plot=True, out_type='MAINT', scenario='BASE'):
    # Scenario is only used when out_type = 'All'

    ## For Debugging
    #print(f' - PSARunID: [{PSARunID}]\n - directory: [{directory}]\n - out_type: [{out_type}]')

    ## JPO: check the hierarchy of the if statements
  
    ## JPO: what is this check for? if we are passing an excel file then no need to look for a directory? but then we need this variable further down
    ## TB: This check is to provide the ability in the UI to pass a different MAINT File than the one in a PSARunID folder, for instance, if you are modifying
    ## a maint file from within the data/input folder.
    directory = directory.replace('/','\\')
    
    if directory[-5:] != '.xlsx':
        c_dict = {'MAINT':'#E69646','CONT':'#E64646'}
        config_dict = read_rtp(PSARunID, directory)
        directory = default_directory(PSARunID, directory)
        save_directory = directory
        if out_type == 'MAINT':
            directory = f'{directory}{const.strINfolder}{config_dict[const.strMaintFile]}'
            thing='OUTAGE'
        elif out_type == 'EVENT':
            directory = f'{directory}{const.strINfolder}{config_dict[const.strEventFile]}'
            c_dict = {'BASE':'#34D05C','MAINT':'#E69646','CONT':'#E64646','MAINT_CONT':'#f51197'}
            thing='SERVICE'
        elif out_type == 'CONT':
            directory = f'{directory}{const.strINfolder}{config_dict[const.strContFile]}'
            thing='OUTAGE'
        elif out_type == 'SWITCH':
            directory = f'{directory}{const.strINfolder}{config_dict[const.strSwitchFile]}'
            c_dict = {'BASE-OPEN':'#6aa84f','BASE-CLOSED':'#d9ead3',
                      'MAINT-OPEN':'#e69138','MAINT-CLOSED':'#fce5cd',
                      'CONT-OPEN':'#cc0000','CONT-CLOSED':'#f4cccc',
                      'MAINT_CONT-OPEN':'#a64d79','MAINT_CONT-CLOSED':'#ead1dc'}
            thing='SERVICE'
        elif out_type == 'MAINT_CONT':
            directory = [f'{directory}{const.strINfolder}{config_dict[const.strMaintFile]}',
                         f'{directory}{const.strINfolder}{config_dict[const.strContFile]}']
            thing=['OUTAGE','OUTAGE']
        elif out_type == 'ALL':
            directory = [f'{directory}{const.strINfolder}{config_dict[const.strMaintFile]}',
                         f'{directory}{const.strINfolder}{config_dict[const.strContFile]}',
                         f'{directory}{const.strINfolder}{config_dict[const.strEventFile]}',
                         f'{directory}{const.strINfolder}{config_dict[const.strSwitchFile]}'] # Uncomment this once config_files have the SWITCH = TRUE/FALSE line + file_loc line
            thing=['OUTAGE','OUTAGE','SERVICE','SERVICE'] # For the names of the columns in the files respectively
        else:
            raise ValueError(f'out_type can only be MAINT, CONT, MAINT_CONT, EVENT, SWITCH. Not {out_type}')
        
    ## TB: if passed a file rather than folder, check if it is Event file or default to it being Maint/Cont file (only matters for determining the column headers)
    ## (this is the bit of code that was missing earlier)
    elif 'EVENT' in directory.upper():
        thing = 'SERVICE'
        c_dict = {'BASE':'#34D05C','MAINT':'#E69646','CONT':'#E64646','MAINT_CONT':'#f51197'}
        save_directory = ('\\').join(directory.split('\\')[:-1])
    elif 'SWITCH' in directory.upper():
        thing = 'SERVICE'
        c_dict = {'BASE-OPEN':'#6aa84f','BASE-CLOSED':'#d9ead3',
                  'MAINT-OPEN':'#e69138','MAINT-CLOSED':'#fce5cd',
                  'CONT-OPEN':'#cc0000','CONT-CLOSED':'#f4cccc',
                  'MAINT_CONT-OPEN':'#a64d79','MAINT_CONT-CLOSED':'#ead1dc'}
        save_directory = ('\\').join(directory.split('\\')[:-1])
    else:
        thing = 'OUTAGE'
        c_dict = {'MAINT':'#E69646','CONT':'#E64646'}
        save_directory = ('\\').join(directory.split('\\')[:-1])

    #print(f'Creating [{out_type}] file in [{directory}]')
    df = []
    if type(directory) == list:
        c_dict = {'MAINT':'#E69646','CONT':'#E64646','EVENT':'#34D05C','OPN-SWITCH':'#f51197','CLS-SWITCH':'#e1d6dc'}
        for idx, filepath in enumerate(directory):
            try:
                df1 = pd.read_excel(filepath, converters={f'START_{thing[idx]}': pd.to_datetime,
                                            f'END_{thing[idx]}': pd.to_datetime})
            except:
                continue
            
            if 'SCENARIO' in list(df1) and scenario != 'All':
                df1 = df1[df1['SCENARIO'] == scenario]

            if thing[idx] != 'OUTAGE':
                df1 = df1.rename(columns={f'START_{thing[idx]}':'START_OUTAGE',f'END_{thing[idx]}':'END_OUTAGE'})
                if 'ACTION' in list(df1):
                    df1['colour'] = [c_dict['OPN-SWITCH'] if row['ACTION'] == 'OPEN' else c_dict['CLS-SWITCH'] for idx, row in df1.iterrows()]
                else:
                    df1['colour'] = list(c_dict.values())[idx]
                
            else:
                df1['colour'] = list(c_dict.values())[idx]
            try:
                df = pd.concat([df, df1])
                df.reset_index(inplace=True, drop=True)
            except:
                df = df1.copy()        
        thing = 'OUTAGE'
    else:
        df = pd.read_excel(directory, converters={f'START_{thing}': pd.to_datetime,
                                        f'END_{thing}': pd.to_datetime})
        df['FILE'] = out_type
        df['colour'] = df.apply(colour, axis=1)
        directory = ('\\').join(directory.split('/'))

    if not isinstance(df, pd.DataFrame):
        raise FileNotFoundError(f'None of the files exist!')
    
    df.fillna(False, inplace=True)
    if thing == 'SERVICE' or out_type == 'ALL':
        df['ASSET_ID'] = [f'{get_initials(scenario_id)}-{asset_id}' if bool(scenario_id) else asset_id for asset_id, scenario_id in zip(df['ASSET_ID'],df['SCENARIO'])]

    if start_time == 'now':
        start_time_now = datetime.now()
        start_time = datetime(*start_time_now.timetuple()[:3])
        start_time_start = start_time_now.timetuple()[3]*2 # This timetuple takes the hour value for the start_time value and returns it. Multipled by 2 to get half-hour.
    else:
        start_time_now = datetime.strptime(PSARunID.split('_')[1], '%Y-%m-%d-%H-%M')
        start_time = datetime(*start_time_now.timetuple()[:3])
        start_time_start = start_time_now.timetuple()[3]*2 

    initials = get_initials(out_type)

    df['start'] = (df[f'START_{thing}']-start_time).dt.days*48 + [df.loc[i,f'START_{thing}'].timetuple()[3]*2 for i in range(len(df))]
    df['start'] = [value if value > 0 else 0 for value in list(df['start'])]
    df['hh'] = (df[f'END_{thing}']-start_time).dt.days*48 + [df.loc[i,f'END_{thing}'].timetuple()[3]*2 for i in range(len(df))] - df['start']
    df['hh'] = [value if start+value < (48*no_of_days) else (48*no_of_days)-start for start, value in zip(list(df['start']), list(df['hh']))]
    df.sort_values(by='start', ascending=True, inplace=True)

    end_time = start_time + timedelta(days=no_of_days)
    
    values = []
    for i in df.index:
        start_relevance = is_time_between(start_time, end_time, df.loc[i,f'START_{thing}'])
        end_relevance = is_time_between(start_time, end_time, df.loc[i,f'END_{thing}'])
        values.append(start_relevance or end_relevance)

    df['RELEVANT'] = values
    df = df[df['RELEVANT'] == True]

    fake_PSARunID = f'PSA-{datetime.strftime(start_time_now, "%Y-%m-%d-%H-%M")}'
    dates_dict = convert_to_dates(fake_PSARunID, range(0,no_of_days))
    labels = list(dates_dict.values())

    end_date = list(dates_dict.values())[-1].split(' ')[1]
    dt_object = datetime.strptime(end_date, '%d/%m/%y') + timedelta(days = 1)
    labels.append(dt_object.strftime('%A')[:3] + ' ' + dt_object.strftime('%d/%m/%y'))

    xticks = [0]

    for i in range(0,no_of_days):
        xticks.append(48 * (i + 1))

    final_dict = dict(zip(xticks, labels))
    
    if plot:
        fig, ax = plt.subplots(1, figsize=(16,6))
        df.reset_index(inplace=True)
        
        i = 0
        used_assets = []
        for idx, row in df.iterrows():
            if row['ASSET_ID'] not in used_assets:
                txt = row['ASSET_ID']
                if 'P_DISPATCH_KW' in list(df): 
                    if not isinstance(row['P_DISPATCH_KW'], bool):
                        txt = f'{row["ASSET_ID"]}\n{row["P_DISPATCH_KW"]} kW'

                plt.text(row['start']+(row['hh']//2), idx-i, txt, va='center', ha='center', alpha=0.8)
                used_assets.append(row['ASSET_ID'])
            else:
                i += 1

        x = len(used_assets)-1
        y, z = 0, 0
        for p, date in enumerate(list(final_dict.values())[:-1]):
            auction = auction_relevance(p, start_time)
            if auction == 'Day ahead' and y == 0:
                plt.fill_between(x=(p*48,(p+1)*48), y1=len(df)+1, y2=-1, facecolor='blue', alpha=0.05)
                plt.text(p*48 + 24, 0, '- Day -\nAhead', color='blue', alpha=0.25, ha='center', va='center', fontsize='large')
                y += 1
            elif auction == 'Week ahead':
                z += 1
                if z == 1:
                    plt.fill_between(x=(p*48,(p+1)*48), y1=len(df)+1, y2=-1, facecolor='blue', alpha=0.05)
                    plt.text(p*48 + 24, x, '- Week -\nAhead', color='blue', alpha=0.25, ha='center', va='center', fontsize='large')
                elif z <= 5:
                    plt.fill_between(x=(p*48,(p+1)*48), y1=len(df)+1, y2=-1, facecolor='blue', alpha=0.05)
                    plt.text(p*48 + 24, x, '--->', color='blue', alpha=0.25, horizontalalignment='center', fontsize='large')    

        legend_elements = [Patch(facecolor=c_dict[i], label=i) for i in c_dict]
        ncols = 2 if len(legend_elements) > 6 else 1
        plt.legend(handles=legend_elements, ncol=ncols)
        plt.axvline(start_time_start, color='blue', alpha=0.8)

        ax.barh(df['ASSET_ID'], df['hh'], left=df['start'], color=df['colour'])

        ax.set_xticks(list(final_dict.keys()), list(final_dict.values()), rotation = 'vertical')
        ax.set_yticks([])
        ax.grid(axis = 'x')

        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_visible(False)
        ax.spines['left'].set_position(('outward', 10))
        ax.spines['top'].set_visible(False)

        ax.set_title(f'{fake_PSARunID} | {initials} {thing}S From Inputs')
        if save:
            return save_result(ax, f'{initials} {thing}_MAPPING', save_directory, new_folder=True)
        else:
            return ax
    else:
        return df, final_dict
    
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''The same as plot_gantt_outages but takes file names explicitly rather than finding them itself'''
def plot_gantt_outages_by_file(event_file='', maint_file='', cont_file='', switch_file='', no_of_days=11, plot=True, save=False): # Add switch file for when switch files are added
    config_dict = ut.PSA_SND_read_config(const.strConfigFile)
    
    c_dict = {'Cont':'#E64646','Maint':'#E69646','Event':'#34D05C','OPN-SWITCH':'#f51197','CLS-SWITCH':'#e1d6dc'}
    df_dict = {}
    df = []
    if os.path.isfile(event_file):
        dfe = pd.read_excel(event_file, converters={'START_SERVICE': pd.to_datetime,
                                             'END_SERVICE': pd.to_datetime})
        dfe = dfe.rename(columns={'START_SERVICE': 'START_OUTAGE', 'END_SERVICE':'END_OUTAGE'})
        dfe['colour'] = c_dict['Event']
        dfe['order'] = 1
        df_dict['Event'] = dfe
    if os.path.isfile(maint_file):
        dfm = pd.read_excel(maint_file, converters={'START_OUTAGE': pd.to_datetime,
                                             'END_OUTAGE': pd.to_datetime})
        dfm['SCENARIO'] = 'MAINT'
        dfm['colour'] = c_dict['Maint']
        dfm['order'] = 2
        df_dict['Maint'] = dfm
    if os.path.isfile(cont_file):
        dfc = pd.read_excel(cont_file, converters={'START_OUTAGE': pd.to_datetime,
                                             'END_OUTAGE': pd.to_datetime})
        dfc['SCENARIO'] = 'CONT'
        dfc['colour'] = c_dict['Cont']
        dfc['order'] = 3
        df_dict['Cont'] = dfc
    if os.path.isfile(switch_file):
        dfs = pd.read_excel(switch_file, converters={'START_SERVICE': pd.to_datetime,
                                             'END_SERVICE': pd.to_datetime})
        dfs = dfs.rename(columns={'START_SERVICE': 'START_OUTAGE', 'END_SERVICE':'END_OUTAGE'})
        dfs['colour'] = [c_dict['OPN-SWITCH'] if row['ACTION'] == 'OPEN' else c_dict['CLS-SWITCH'] for idx, row in dfs.iterrows()]
        dfs['order'] = 0
        df_dict['Switch'] = dfs
    if type(df_dict) == list:
        raise FileNotFoundError(f'None of the files provided exist!\nevent_file: {event_file}\nmaint_file: {maint_file}\ncont_file: {cont_file}\nswitch_file: {switch_file}')
    
    df_list = list(df_dict.values())
    for dfx in df_list:
        try:
            df = pd.concat([df, dfx])
        except:
            df = dfx.copy()
    
    if len(df) == 0:
        raise KeyError('Nothing to Show')
    
    df.reset_index(drop=True, inplace=True)
    df.fillna(False, inplace=True)

    if 'SCENARIO' in list(df):
        df['ASSET_ID'] = [f'{get_initials(scenario_id)}-{asset_id}' if bool(scenario_id) else asset_id for asset_id, scenario_id in zip(df['ASSET_ID'],df['SCENARIO'])]
    
    start_time_now = datetime.now()
    start_time_start = start_time_now.timetuple()[3]*2
    start_time = datetime(*start_time_now.timetuple()[:3])
    
    df['start'] = (df[f'START_OUTAGE']-start_time).dt.days*48 + [df.loc[i,f'START_OUTAGE'].timetuple()[3]*2 for i in range(len(df))]
    df['start'] = [value if value > 0 else 0 for value in list(df['start'])]
    df['hh'] = (df[f'END_OUTAGE']-start_time).dt.days*48 + [df.loc[i,f'END_OUTAGE'].timetuple()[3]*2 for i in range(len(df))] - df['start']
    df['hh'] = [value if start+value < (48*no_of_days) else (48*no_of_days)-start for start, value in zip(list(df['start']), list(df['hh']))]
    df.sort_values(by=['SCENARIO','order'], ascending=True, inplace=True) # Changed from sort_values(by=['SCENARIO','start'])
    
    end_time = start_time + timedelta(days=no_of_days)
    
    values = []
    for i in df.index:
        start_relevance = is_time_between(start_time, end_time, df.loc[i,f'START_OUTAGE'])
        end_relevance = is_time_between(start_time, end_time, df.loc[i,f'END_OUTAGE'])
        values.append(start_relevance or end_relevance)

    df['RELEVANT'] = values
    df = df[df['RELEVANT'] == True]

    fake_PSARunID = f'PSA-{datetime.strftime(start_time_now, "%Y-%m-%d-%H-%M")}'
    dates_dict = convert_to_dates(fake_PSARunID, range(0,no_of_days))
    labels = list(dates_dict.values())

    end_date = list(dates_dict.values())[-1].split(' ')[1]
    dt_object = datetime.strptime(end_date, '%d/%m/%y') + timedelta(days = 1)
    labels.append(dt_object.strftime('%A')[:3] + ' ' + dt_object.strftime('%d/%m/%y'))

    xticks = [0]

    for i in range(0,no_of_days):
        xticks.append(48 * (i + 1))

    final_dict = dict(zip(xticks, labels))
    assets_no = len(df['ASSET_ID'])
    fig_height = 6 if 6 > assets_no//2 else assets_no//2 if assets_no//2 < 18 else 18
    if plot:
        fig, ax = plt.subplots(1, figsize=(16,fig_height))
        df.reset_index(inplace=True)

        i = 0
        used_assets = []
        for idx, row in df.iterrows():
            if row['ASSET_ID'] not in used_assets: 
                txt = row['ASSET_ID'] if row['colour'] != c_dict['Event'] else f'{row["ASSET_ID"]}\n{row["P_DISPATCH_KW"]} kW'
                plt.text(row['start']+(row['hh']//2), idx-i, txt, va='center', ha='center', alpha=0.8)
                used_assets.append(row['ASSET_ID'])
            else:
                i += 1

        x = len(used_assets)-1
        y, z = 0, 0
        for p, date in enumerate(list(final_dict.values())[:-1]):

            auction = auction_relevance(p, start_time)
            if auction == 'Day ahead' and y == 0:
                plt.fill_between(x=(p*48,(p+1)*48), y1=len(df)+1, y2=-1, facecolor='blue', alpha=0.05)
                plt.text(p*48 + 24, 0, '- Day -\nAhead', color='blue', alpha=0.25, ha='center', va='center', fontsize='large')
                y += 1
            elif auction == 'Week ahead':
                z += 1
                if z == 1:
                    plt.fill_between(x=(p*48,(p+1)*48), y1=len(df)+1, y2=-1, facecolor='blue', alpha=0.05)
                    plt.text(p*48 + 24, x, '- Week -\nAhead', color='blue', alpha=0.25, ha='center', va='center', fontsize='large')
                elif z <= 5:
                    plt.fill_between(x=(p*48,(p+1)*48), y1=len(df)+1, y2=-1, facecolor='blue', alpha=0.05)
                    plt.text(p*48 + 24, x, '--->', color='blue', alpha=0.25, horizontalalignment='center', fontsize='large')    

        y_dict, y_majors = {}, []
        y1 = -0.5
        for scenario in list(set(df['SCENARIO'])):
            q = len(set(df[df['SCENARIO'] == scenario]['ASSET_ID']))
            y1 += q
            y2 = y1 - q//2
            y_majors.append(y1)
            y_dict[y2] = scenario

        legend_elements = [Patch(facecolor=c_dict[i], label=i) for i in c_dict]
        plt.legend(handles=legend_elements)
        plt.axvline(start_time_start, color='blue', alpha=0.8)

        ax.barh(df['ASSET_ID'], df['hh'], left=df['start'], color=df['colour'])    

        ax.set_xticks(list(final_dict.keys()), list(final_dict.values()), rotation = 'vertical')
        ax.set_yticks(list(y_dict.keys()), list(y_dict.values()), va = 'center')

        ax.yaxis.set_major_locator(mpl.ticker.FixedLocator(y_majors))
        ax.yaxis.set_minor_locator(mpl.ticker.FixedLocator(list(y_dict.keys())))
        ax.yaxis.set_major_formatter(mpl.ticker.NullFormatter())
        ax.set_yticklabels(list(y_dict.values()), minor=True, color='blue', weight='demi')

        ax.grid(axis = 'both')

        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_visible(True)
        ax.spines['left'].set_position(('outward', 10))
        ax.spines['top'].set_visible(False)

        #fake_PSARunID = fake_PSARunID.replace('00-00','xx-xx')
        ax.set_title(f'{fake_PSARunID} | All Asset Activity Data')
    else:
        return df #, final_dict for debugging
    
    if save:
        save_directory = f'{config_dict[const.strWorkingFolder]}{const.strDataInput[:-1]}'
        return save_result(ax, f'Latest_PSA_Run_ALL_OUTAGE_MAPPING', save_directory, new_folder=True)
    else:
        return ax    

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function works identically to plot_gantt_outages but can only plot one asset, and takes its input from the asset outage tables than a maint file'''
def plot_gantt_outages_debug(PSARunID, directory, out_type='MAINT', plot=True, save=False):
    # When out_type == 'MAINT_CONT' the behaviour of this function is different to plot outages from inputs as there is no way of knowing where the outage tables files have come from
    # A workaround could be to read both the maint and cont outage tables seperately and colour accordingly. 

    directory = default_directory(PSARunID, directory)
    initials = get_initials(out_type)
    data_df = pd.DataFrame(columns=['asset','colour','start','hh'])
    asset_file = ''

    if out_type == 'SWITCH':
        c_dict = {'OPN-SWITCH':'#f51197','CLS-SWITCH':'#e1d6dc'}
        asset_outage_tables = directory
        folders = [const.strSwitchFiles.strip('\\')]
        switch = True
    else:
        c_dict = {'MAINT':'#E69646','CONT':'#E64646','MAINT_CONT':'#a64d79'}
        asset_outage_tables = f'{directory}{const.strINfolder}{initials}-ASSET_OUTAGE_TABLES'
        folders = os.listdir(asset_outage_tables)
        switch = False

    for folder in folders:
        key = ['CLS-SWITCH', 'OPN-SWITCH'] if switch else [out_type, out_type]
        for file in os.listdir(f'{asset_outage_tables}\\{folder}'): # Read each asset in both lines and transformers
            asset_file = f'{asset_outage_tables}\\{folder}\\{file}'
            df = pd.read_excel(asset_file, index_col = 0)
            df = pd.melt(df, var_name='Day',value_name='Out_bool')
            df_diff = df.diff()
            df_diff.loc[0, 'Out_bool'] = df.loc[0, 'Out_bool']
            df_diff.loc[len(df)] = [0,df.loc[len(df)-1, 'Out_bool']]
            df_diff = df_diff[df_diff['Out_bool']!=0]
            df_diff.reset_index(inplace=True)
            if len(df_diff) >= 2:
                asset = file.removesuffix('.xlsx')
                if switch:
                    data_df.loc[len(data_df)] = [asset,c_dict[key[0]],0,df.index.max()+1] # This starts at 0 hence max = 527 on 11 day run hence the +1
                for i in range(0, len(df_diff), 2):
                    start = df_diff.loc[i,'index']
                    hh = df_diff.loc[i+1,'index'] - df_diff.loc[i,'index']
                    data_df.loc[len(data_df)] = [asset,c_dict[key[1]],start,hh]

    if data_df.shape[0] == 0:
        raise FileNotFoundError(f'No assets are ever out of service in {initials}-ASSET_OUTAGE_TABLES')
    
    no_of_days = int(round(df.index.max()/48+0.5,0))
    df = data_df.sort_values(by='start', ignore_index=True)

    start_time_now = datetime.strptime(PSARunID.split('_')[1], '%Y-%m-%d-%H-%M')
    start_time_start = start_time_now.timetuple()[3]*2
    start_time = datetime(*start_time_now.timetuple()[:3])
    
    fake_PSARunID = f'PSA-{datetime.strftime(start_time, "%Y-%m-%d-%H-%M")}'
    dates_dict = convert_to_dates(fake_PSARunID, range(0,no_of_days))
    labels = list(dates_dict.values())

    end_date = list(dates_dict.values())[-1].split(' ')[1]
    dt_object = datetime.strptime(end_date, '%d/%m/%y') + timedelta(days = 1)
    labels.append(dt_object.strftime('%A')[:3] + ' ' + dt_object.strftime('%d/%m/%y'))

    xticks = [0]

    for i in range(0,no_of_days):
        xticks.append(48 * (i + 1))

    final_dict = dict(zip(xticks, labels))
    used_assets = []
  
    if plot:
        fig, ax = plt.subplots(1, figsize=(16,6))

        for idx, start in enumerate(df['start']):
            txt = df.loc[idx,'asset']
            if txt not in used_assets:
                plt.text(df['start'][idx]+(df['hh'][idx]//2), idx, txt, va='center', ha='center', alpha=0.8)
                used_assets.append(txt)

        x = len(df['asset'].unique())-1
        y, z = 0, 0
        for p, date in enumerate(list(final_dict.values())[:-1]):
            auction = auction_relevance(p, start_time)
            if auction == 'Day ahead' and y == 0:
                plt.fill_between(x=(p*48,(p+1)*48), y1=len(df)+1, y2=-1, facecolor='blue', alpha=0.05)
                plt.text(p*48 + 24, 0, '- Day -\nAhead', color='blue', alpha=0.25, va='center', ha='center', fontsize='large')
                y += 1
            elif auction == 'Week ahead':
                z += 1
                if z == 1:
                    plt.fill_between(x=(p*48,(p+1)*48), y1=len(df)+1, y2=-1, facecolor='blue', alpha=0.05)
                    plt.text(p*48 + 24, x, '- Week -\nAhead', color='blue', alpha=0.25, va='center', ha='center', fontsize='large')
                elif z <= 5:
                    plt.fill_between(x=(p*48,(p+1)*48), y1=len(df)+1, y2=-1, facecolor='blue', alpha=0.05)
                    plt.text(p*48 + 24, x, '--->', color='blue', alpha=0.25, horizontalalignment='center', fontsize='large')    

        legend_elements = [Patch(facecolor=c_dict[i], label=i) for i in c_dict]
        plt.legend(handles=legend_elements)
        plt.axvline(start_time_start, color='blue', alpha=0.8)

        ax.barh(df['asset'], df['hh'], left=df['start'], color=df['colour'])    

        ax.set_xticks(list(final_dict.keys()), list(final_dict.values()), rotation = 'vertical')
        ax.set_yticks([])
        ax.grid(axis = 'x')

        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_visible(False)
        ax.spines['left'].set_position(('outward', 10))
        ax.spines['top'].set_visible(False)

        ax.set_title(f'{fake_PSARunID} | {initials} OUTAGES From Tables')
    else:
        return df #, final_dict # For Debugging

    if save:
        #save_directory = ('\\').join(directory.split('\\')[:-1])
        return save_result(ax, f'{initials} OUTAGE_TABLES_MAPPING', directory, new_folder=True)
    else:
        return ax
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function checks if a certain check time takes place between a start and and end time'''
def is_time_between(begin_time, end_time, check_time=None):
    check_time = check_time or datetime.now()
    truth = check_time >= begin_time and check_time <= end_time
    return truth

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Potentially Useful Function that plots a dataframe on a matplotlib graph which can be saved as png rather than xlsx'''
def render_mpl_table(data, col_width=3.0, row_height=0.625, font_size=14,
                     header_color='#40466e', row_colors=['#f1f1f2', 'w'], edge_color='w',
                     bbox=[0, 0, 1, 1], header_columns=0,
                     ax=None, **kwargs):
    if ax is None:
        size = (np.array(data.shape[::-1]) + np.array([0, 1])) * np.array([col_width, row_height])
        fig, ax = plt.subplots(figsize=size)
        ax.axis('off')
    mpl_table = ax.table(cellText=data.values, bbox=bbox, colLabels=data.columns, **kwargs)
    mpl_table.auto_set_font_size(False)
    mpl_table.set_fontsize(font_size)

    for k, cell in mpl_table._cells.items():
        cell.set_edgecolor(edge_color)
        if k[0] == 0 or k[1] < header_columns:
            cell.set_text_props(weight='bold', color='w')
            cell.set_facecolor(header_color)
        else:
            cell.set_facecolor(row_colors[k[0]%len(row_colors) ])
    return ax.get_figure(), ax

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This function reads a Nerda file and returns only the rows where state_before != state_after'''
def get_nerda(PSARunID, directory):
    directory = default_directory(PSARunID, directory)
    file = f'{directory}{const.strNeRDAfolder}\\switch_status.xlsx'
    df_dict = pd.read_excel(file, sheet_name=['Switch_before','Switch_after','Coupler_before','Coupler_after'], index_col=0, usecols=[0,1,2,3,4,5,6,7,9])
    # Use cols ['Unnamed','Name','Grid','Folder','Terminal','Network_element','Network_element_type','Closed','Type']
    
    # Switches
    df_switch = df_dict['Switch_before'].join(df_dict['Switch_after']['Closed'], on='Closed', lsuffix='_before', rsuffix='_after')
    df_switch = df_switch[df_switch['Closed_before']!=df_switch['Closed_after']]
    
    df_switch[['Closed_before','Closed_after']] = df_switch[['Closed_before','Closed_after']].replace({True: 'CLOSED',False: 'OPEN'})
    df_switch.rename(columns={'Closed_before':'State_before','Closed_after':'State_after'},inplace=True)

    # Couplers
    df_coupler = df_dict['Coupler_before'].join(df_dict['Coupler_after']['Closed'], on='Closed', lsuffix='_before', rsuffix='_after')
    df_coupler = df_coupler[df_coupler['Closed_before']!=df_coupler['Closed_after']]
    
    df_coupler[['Closed_before','Closed_after']] = df_coupler[['Closed_before','Closed_after']].replace({True: 'CLOSED',False: 'OPEN'})
    df_coupler.rename(columns={'Closed_before':'State_before','Closed_after':'State_after'},inplace=True)

    # Return Dictionary
    new_dict = {'Switches': df_switch,
                'Couplers': df_coupler}

    return new_dict

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''For file names ending in -V(i).xlsx, this function finds the file with the highest version number and returns it'''
def latest_version(file):
    file_list = file.split('.')
    if file_list != [file]:
        file = file_list[0]
        if not file[-1].isnumeric():
            file = f'{file}-V0'
        extension = file_list[1]
        i = int(file[-1])
        if not os.path.isfile(f'{file}.{extension}'):
            raise FileNotFoundError(f'File does not exist: {file}.{extension}')
        while os.path.isfile(f'{file}.{extension}'):
            i += 1
            file = f'{file[:-1]}{i}' 
        return f'{file[:-1]}{i-1}.{extension}'
    else:
        if not file[-1].isnumeric():
            file = f'{file}\\V0'
        i = int(file[-1])
        if not os.path.isdir(file): 
            raise FileNotFoundError(f'Folder does not exist: {file}')
        while os.path.isdir(file):
            i += 1
            file = f'{file[:-1]}{i}'
        return f'{file[:-1]}{i-1}'

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Returns the batch number of the highest SF file in a folder'''
def latest_batch(PSARunID, directory=None):
    directory = default_directory(PSARunID, directory)
    sf_file = f'{directory}\\{PSARunID}{const.strPSAResponses}.xlsx'
    max_sf = latest_version(sf_file)
    batch = max_sf.removesuffix('.xlsx').split('-')[-1][1:]
    return int(batch)

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''For file names ending in -V(d)-(i).xlsx, this function finds the file with the highest version number and returns it'''
def batch_latest_version(file, PSARunID, directory=None):
    batch = latest_batch(PSARunID, directory)

    file_list = file.split('.')
    if file_list != [file]:
        file_list[0] = f'{file_list[0]}-V{batch}-0'
        file = ('.').join(file_list)
    else:
        file = f'{file}\\V{batch}-0'

    # This breaks if passed a file/fldr with V0 (or some variation of) already on the end. It should be blank.
    return latest_version(file)

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This Function returns the batch and version number of a file either as a string or two integer values'''
def get_batch_version_number(file, concatenate=True):
    batch, version = file.removesuffix('.xlsx').split('V')[1].split('-')
    if concatenate:
        return f'{batch}-{version}'
    else:
        return int(batch), int(version)

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This Function reads the conresfiles of a PSARunID and retuns either a bar chart or raw data'''

def read_conres(PSARunID, directory, version_number='latest', save=False):
    directory = default_directory(PSARunID, directory=directory)
    if version_number == 'latest':
        fldr = f'{directory}\\TEMP_CR_FILES'
        files = batch_latest_version(fldr, PSARunID, directory) # changed from latest version to batch latest version
        version_number = get_batch_version_number(files)
    else:
        files = f'{directory}\\TEMP_CR_FILES\\V{version_number}'

    dfd = read_conres_debug(PSARunID, directory, version_number=version_number, save=False)
    valid_req_ids = dfd.index.get_level_values(2).drop_duplicates().to_list()
    conres = [file for file in os.listdir(files) if file[:6] == 'ConRes' and int(file.split('_')[-2]) in valid_req_ids]

    df = pd.DataFrame(columns=['req_id','resp_id','constrained_pf_id'], index=range(len(conres)))
    df_old = []
    for i, file in enumerate(conres):
        file_trim = file[11:-4].split('_')

        df.loc[i] = [file_trim[-2], file_trim[-1], f'{("_").join(file_trim[:-2])}']
        df.loc[i+1] = [file_trim[-2], file_trim[-1], f'{("_").join(file_trim[:-2])}']

        df_new = pd.read_csv(os.path.join(files, file), usecols=['Unnamed: 0',
                                                                     'loading_threshold_pct',
                                                                     'Loading','tot_offered_power_kw',
                                                                     'P_req', # Changed res_flex_power_kw -> P-req
                                                                     'tot_responses'])
        df_new['resp_id'] = file_trim[-1]
        df_new['req_id']  = file_trim[-2]
        df_new['constrained_pf_id'] = f'{("_").join(file_trim[:-2])}'
        try:
            df_old = pd.concat([df_new, df_old], ignore_index=True)
        except:
            df_old = df_new.copy()

    df_old['req_id']  = df_old['req_id'].astype(int)
    df_old['resp_id'] = df_old['resp_id'].astype(int)
    df_old.columns = ['calc_step' if x=='Unnamed: 0' else x for x in df_old.columns]

    req_to_start = dict(zip(dfd.index.get_level_values(2),dfd.index.get_level_values(3)))
    resp_to_flex = dict(zip(dfd.index.get_level_values(4),dfd.flex_pf_id))
    req_to_scenario = dict(zip(dfd.index.get_level_values(2),dfd.index.get_level_values(0)))

    df_old['start_time'] = [req_to_start[key] for key in df_old['req_id']]
    df_old['scenario'] = [req_to_scenario[key] for key in df_old['req_id']]
    df_old['flex_pf_id'] = [resp_to_flex[key] if df_old.loc[i, 'tot_responses'] == 1 else 'Multiple' for i, key in enumerate(df_old['resp_id'])]
    df_old['overload'] = [True if loading > threshold else False for loading, threshold in zip(df_old['Loading'], df_old['loading_threshold_pct'])]

    dfs = read_sf_by_batch(PSARunID, directory, save=False)
    dfs = dfs.reset_index().set_index('req_id')['batch'].to_frame()
    df_old['batch'] = [dfs.loc[key,'batch'] for key in df_old['req_id']]

    final_df = df_old.set_index(['scenario','constrained_pf_id','req_id','start_time','resp_id']) #,'calc_step'
    final_df = final_df.sort_index()

    if save:
        return save_result(final_df, f'ConRes_TEMP-FILE_DATA_V{version_number}', directory)
    else:
        return final_df

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''uses read_conres to plot a bar chart'''
def plot_conres(PSARunID, directory, asset=None, scenario=None, version_number='latest', plot=True, save=False):
    directory = default_directory(PSARunID, directory)

    if version_number == 'latest':
        fldr = f'{directory}\\TEMP_CR_FILES'
        files = batch_latest_version(fldr, PSARunID, directory) # changed from latest version to batch latest version
        version_number = get_batch_version_number(files)
    else:
        files = f'{directory}\\TEMP_CR_FILES\\V{version_number}'
    
    new_df = read_conres(PSARunID, directory, version_number, save=False)
    if not asset:
        asset = new_df.index.levels[1][0]
    if not scenario:
        scenario = new_df.index.levels[0][0]
    
    filt_df = new_df.loc[scenario, asset]

    fin_df = filt_df.loc[(filt_df['calc_step'] == 'final'),
                         ['P_req','tot_offered_power_kw','flex_pf_id','Loading']]
    fin_df = fin_df.reindex(fin_df['Loading'].sort_values(ascending=True).index)
    fin_df = fin_df.reset_index().drop_duplicates('req_id').set_index(['req_id','resp_id'])
    fin_df = fin_df.sort_values(by='start_time')
    
    init_df = filt_df.loc[(filt_df['calc_step'] == 'initial'),
                         ['P_req','tot_offered_power_kw','flex_pf_id','Loading']]
    init_df = init_df.reset_index().set_index(['req_id','resp_id'])
    init_df = init_df.reindex_like(fin_df)
    
    final_df = fin_df.copy()
    final_df['init_flex_req_power_kw'] = init_df['P_req']
    final_df['init_Loading'] = init_df['Loading']

    if plot:
        final_df.reset_index(inplace=True)
        si = get_initials(scenario)
        head=len(final_df)

        x1 = list(final_df['req_id'])
        x2 = np.arange(head)

        y1 = list(final_df['init_flex_req_power_kw'])
        y2 = list(final_df['tot_offered_power_kw'])
        y3 = list(final_df['P_req'])

        y4 = list(final_df['init_Loading'])
        y5 = list(final_df['Loading'])
        
        fig_width = 5 if head < 6 else head if head < 11 else head//2
        fig, ax1 = plt.subplots(figsize = (fig_width, 4))

        ax1.bar(x2-0.2, y1[:head], width=0.2, color='red', alpha=0.3)
        ax1.bar(x2, y2[:head], width=0.2, color='green', alpha=0.3)
        ax1.bar(x2+0.2, y3[:head], width=0.2, color='blue', alpha=0.3)
        ax1.set_xlabel('Req ID (Resp ID)')
        ax1.set_ylabel('Required kW')
        ax1.legend(['Initial Req','Offered Power','Residual Req'], loc='lower center', bbox_to_anchor=(0, 1))
        #ax1.grid(axis='x')
        
        threshold = get_asset_threshold(asset, PSARunID, directory)
        ax2 = ax1.twinx()
        ax2.scatter(x2-0.2, y4, marker='^', color='red', alpha=0.3)
        ax2.axhline(y = threshold, color = 'red', label = f'{threshold}% Threshold')
        ax2.scatter(x2+0.2, y5, marker='v', color='blue', alpha=0.3)
        ax2.set_ylabel('% Loading')
        ax2.legend(['Initial Load', f'{threshold}% Threshold', 'Residual Load'], loc='lower center', bbox_to_anchor=(1, 1))

        for idx, row in final_df.iterrows():
            ax2.plot([idx-0.2, idx+0.2], [row['init_Loading'],row['Loading']], color='silver')

        ax1_ylims = ax1.axes.get_ylim()           # Find y-axis limits set by the plotter
        ax1_yratio = ax1_ylims[0] / ax1_ylims[1]  # Calculate ratio of lowest limit to highest limit
        ax2_ylims = ax2.axes.get_ylim()           # Find y-axis limits set by the plotter
        ax2_yratio = ax2_ylims[0] / ax2_ylims[1]  # Calculate ratio of lowest limit to highest limit
        
        yabs_max = abs(max((ax2_ylims[0]-threshold, ax2_ylims[1]-threshold), key=abs))
        if ax1_yratio < ax2_yratio: 
            ax2.set_ylim(bottom = threshold+(yabs_max*ax1_yratio)) #threshold+(yabs_max*ax1_yratio)
        else:
            ax1.set_ylim(bottom = ax1_ylims[1]*ax2_yratio)

        labels = [f'{req_id} ({resp_id})' if head < 11 else f'{req_id}\n({resp_id})' for req_id, resp_id in zip(list(final_df['req_id'])[:head], list(final_df['resp_id'])[:head])]
        plt.xticks(x2, labels)
        plt.title(f'{si}: {asset} (V{version_number})')
        
        if save:
            return save_result(ax1, f'ConRes_V{version_number}_GRAPH', directory)
        else:
            return ax1
    else:
        return final_df

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''read_conres debug is a reading of the candidates responses file and subsequent formatting to be in line with the output of read_conres'''
def read_conres_debug(PSARunID, directory, version_number='latest', save=False):
    directory = default_directory(PSARunID, directory)

    if version_number == 'latest':
        file = f'{directory}\\{PSARunID}{const.strSNDCandidateResponses}.xlsx'
        file = batch_latest_version(file, PSARunID, directory) # changed from latest version to batch latest version
        version_number = get_batch_version_number(file)
    else:
        file = f'{directory}\\{PSARunID}{const.strSNDCandidateResponses}-V{version_number}.xlsx'
    
    df = pd.read_excel(file)
    df = df.set_index(['scenario','constrained_pf_id','req_id','start_time','resp_id'])[['flex_pf_id','required_power_kw','offered_power_kw']]
    df = df.sort_index(level=2, ascending=False)
    if save:
        return save_result(df, f'ConRes_SND_DATA_V{version_number}', directory)
    else:
        return df

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Plot Events P_dispatch'''
def plot_p_dispatch(asset, PSARunID, directory, save=False):
    directory = default_directory(PSARunID, directory=directory)
    
    event_data = f'{directory}{const.strEventFiles}{asset}.xlsx'

    # Read the excel file
    df_line = pd.read_excel(event_data, index_col=0)
    column_names = convert_to_dates(PSARunID, df_line)
    df_line.rename(columns = column_names, inplace = True)

    df_melt = pd.melt(df_line, var_name='Day', value_name='P_Dispatch')
    title = f'{PSARunID} | {asset} Power Dispatch Graph'

    ax = plot_generic_graph(PSARunID, df_melt, title, y_label='P_Dispatch (kW)', colour='darkgreen')

    if save:
        return save_result(ax, f'{asset}_P_DISPATCH_GRAPH', directory)
    else:
        return ax

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Plot Asset data plots all the thresholds used in the asset data file'''
def plot_generic_graph(PSARunID, df, title=None, colour='Blue', linestyle='-', alpha=1, decorate='type_1', y_label=None, axes=None, vary='colour', h_line=False, describer=False):
    date = PSARunID.split('_')[1]
    dt_object = datetime.strptime(date, '%Y-%m-%d-%H-%M')
    possible_decorations = ['type_1','type_2','type_3','type_4']
    # Type 1 Highlights with Green and Yellow background the SPM and SEPM window
    # Type 2 Highlights with blue background the day ahead/week ahead relevence
    # Type 3 is no decoration
    # Type 4 does both type 1 and type 2
    legend_elements = []
    return_legend = False

    if 'Day' != list(df)[0]:
        raise TypeError(f'Data should be in format Day, variable_1, variable_2... etc with HH set as index. Current data in format {(", ").join(list(df))}')
    if decorate not in possible_decorations:
        raise ValueError(f'Decorate can only be {(",").join(possible_decorations)} not {decorate}')

    column_names = df['Day'].drop_duplicates().to_list() # This variable really needs to have its name changed as it's not column names it is day names in the format "AAA DD/MM/YY" (As output by convert_to_dates) 
    columns = list(df)[1:]
    SPM_colour='green'
    SEPM_colour='yellow'

    prop_cycle = plt.rcParams['axes.prop_cycle']
    colours = prop_cycle.by_key()['color']

    if not y_label:
        plot_y_label = list(df)[1]
    else:
        plot_y_label = y_label
    if not title:
        plot_title = f'{PSARunID}: {plot_y_label} Graph'
    else:
        plot_title = title
    if not axes:
        fig, ax = plt.subplots(figsize=(12, 4))
    else:
        ax = axes 
        return_legend = True
    if type(colour) == str:
        colour = [colour]*len(columns)
    elif len(colour) != len(columns):
        raise ValueError(f'Not enough colours ({len(colour)}) for columns ({len(columns)})')
    if type(linestyle) == str:
        linestyle = [linestyle]*len(columns)
    elif len(linestyle) != len(columns):
        raise ValueError(f'Not enough colours ({len(linestyle)}) for columns ({len(columns)})')
        
    for q, value in enumerate(columns):
        s = linestyle[q]
        if len(columns) == 1:
            c = colour[q]
            a = alpha
        elif vary == 'colour':
            c = colours[q]
            a = alpha
        elif vary == 'alpha':
            c = colour[q]
            a = 1.0/(len(columns) - q)
        if decorate == 'type_1' or decorate == 'type_4':
            if len(columns) > 1:
                legend_elements.append(Line2D([0], [0], label=value, linestyle=s, color=c, alpha=a))
            for i, date in enumerate(column_names):
                auction = auction_relevance(i, dt_object)
                df_date = df[df['Day'] == date]
                if auction == 'Day ahead':
                    ax.plot(df_date[value], label = f'Day-ahead | {date}', linestyle='--', color=c, alpha=a)
                elif auction == 'Week ahead':
                    ax.plot(df_date[value], label = f'Week-ahead | {date}', linestyle='-', color=c, alpha=a)
                else:
                    ax.plot(df_date[value], label = f'{auction} | {date}', linestyle=':', color=c, alpha=a)
        elif decorate == 'type_2' or decorate == 'type_3':
            legend_elements.append(Line2D([0], [0], label=value, linestyle=s, color=c, alpha=a))
            ax.plot(df[value], label = date, linestyle=s, color=c, alpha=a)
       
    xticks = [0]
    y1 = max(df.max()[1:])
    y2 = min(df.min()[1:])  
    
    y, z = 0, 0
    if decorate != 'type_3':
        for i in range(len(column_names)):
            xticks.append(48 * (i + 1))
            if decorate == 'type_1' or decorate == 'type_4':
                ax.fill_between(x=(29+(48*i),37+(48*i)), y1=y1, y2=y2, facecolor=SPM_colour, alpha=0.2)
                ax.fill_between(x=(19+(48*i),27+(48*i)), y1=y1, y2=y2, facecolor=SEPM_colour, alpha=0.2)
            elif decorate == 'type_2' or decorate == 'type_4':
                auction = auction_relevance(i, dt_object)
                if auction == 'Day ahead' and y == 0:
                    ax.fill_between(x=(i*48,(i+1)*48), y1=y1, y2=y2, facecolor='blue', alpha=0.05)
                    ax.text(i*48 + 24, y2+((y1-y2)//2), '- Day -\nAhead', color='blue', alpha=0.25, va='center', ha='center', fontsize='medium')
                    y += 1
                elif auction == 'Week ahead':
                    z += 1
                    if z == 1:
                        ax.fill_between(x=(i*48,(i+1)*48), y1=y1, y2=y2, facecolor='blue', alpha=0.05)
                        ax.text(i*48 + 24, y2+((y1-y2)//2), '- Week -\nAhead', color='blue', alpha=0.25, va='center', ha='center', fontsize='medium')
                    elif z <= 5:
                        ax.fill_between(x=(i*48,(i+1)*48), y1=y1, y2=y2, facecolor='blue', alpha=0.05)
                        ax.text(i*48 + 24, y2+((y1-y2)//2), '--->', color='blue', alpha=0.25, horizontalalignment='center', fontsize='medium')    

    labels = list(column_names)
    end_date = list(column_names)[-1].split(' ')[1]
    dt_object = datetime.strptime(end_date, '%d/%m/%y') + timedelta(days = 1)
    labels.append(dt_object.strftime('%A')[:3] + ' ' + dt_object.strftime('%d/%m/%y'))
    
    if len(columns) == 1 and decorate not in ['type_2', 'type_4']:
        legend_elements += [Line2D([0], [0], color=c, linestyle='-', label='Week Ahead Auction'),
                            Line2D([0], [0], color=c, linestyle='--', label='Day Ahead Auction'),
                            Line2D([0], [0], color=c, linestyle=':', label='No Auction Relevance')]
    if decorate in ['type_1', 'type_4']:
        legend_elements.append(Patch(facecolor=SPM_colour, alpha=0.2, label='SPM Window'))
        legend_elements.append(Patch(facecolor=SEPM_colour, alpha=0.2, label='SEPM Window'))
    if not isinstance(h_line, bool):
        ax.axhline(h_line, color='Red')
        legend_elements.append(Line2D([0], [0], color='Red', linestyle='-', label=f'{h_line}% Threshold'))

    ncols = 1 + len(legend_elements)//6

    ax.grid(axis = 'x')
    ax.set_ylabel(plot_y_label)
    ax.set_xticks(xticks, labels, rotation = 'vertical')
    if not return_legend:
        ax.set_title(plot_title, loc='left')
        if ncols < 5:
            ax.legend(bbox_to_anchor=(0., 1.02, 1., .102), handles=legend_elements, loc='lower right', ncol=ncols, borderaxespad=0.)
        else:
            print(f'Too many elements to fit in legend: {len(legend_elements)}')
        if describer and len(columns) == 1:
            # This doesn't like if any of the column names have spaces in them
            keys, values = describe_helper(df, list(df)[1], target_stats=['min','max','mean','std'])
            keys = keys.replace('mean','avg') 
            plt.figtext(0.45, 0.86, keys, {'multialignment':'left'})
            plt.figtext(0.5, 0.86, values, {'multialignment':'right'})
        return ax
    else:
        return legend_elements
    
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Plot Asset data plots all the thresholds used in the asset data file'''
def plot_asset_data(PSARunID, directory, relevant=False, save=False):
    if directory[-5:] != '.xlsx':
        config_dict = read_rtp(PSARunID, directory)
        directory = default_directory(PSARunID, directory)
        save_directory = directory
        asset_data_file = config_dict[const.strAssetData]
        asset_data = f'{directory}{const.strINfolder}{asset_data_file}'
    else:
        directory = directory.replace('\\','/')
        asset_data_file = directory.split('/')[-1][:-5]
        asset_data = directory
        save_directory = ('/').join(directory.split('/')[:-1])

    try:
        df = pd.read_excel(asset_data, usecols=[0,1,2,6])
        df = df.rename(columns={'ASSET_LONG_NAME':'ASSET_ID'})
        if relevant:
            df = df.loc[(df['EVENTS']==True) & (df['ASSET_TYPE']!=const.strAssetGen)]     
    except BaseException as err:
        print(f'Error reading Asset file: [Errno] {err}\n Defaulting')
        df = pd.read_excel(asset_data, usecols=[0,1,2])
        
    c_dict = {const.strAssetLine:'#E64646',const.strAssetTrans:'#E69646',const.strAssetGen:'#34D05C',const.strAssetLoad:'#f51197'}
    df['colour'] = df.apply(colour, axis=1)
    df = df.sort_values(by='ASSET_ID', ascending=False, ignore_index=True) # set by='MAX_LOADING' to sort by max loading
    
    fig_size = len(df)//4 if len(df)//4 > 10 else 5
    fig, ax = plt.subplots(1, figsize=(8,fig_size))

    ax.barh(df['ASSET_ID'],df['MAX_LOADING'],color=df['colour'],alpha=0.75)

    legend_elements = [Patch(facecolor=c_dict[i], alpha=0.75, label=i) for i in c_dict]
    plt.legend(handles=legend_elements)

    for idx, row in df.iterrows():
        txt = row['MAX_LOADING']
        plt.text(row['MAX_LOADING'], idx, txt, va='center', ha='right', alpha=0.8)

    ax.set_ylabel('')
    ax.set_xlabel('Threshold (%)')
    ax.grid(axis = 'x')

    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.spines['left'].set_position(('outward', 10))
    ax.spines['top'].set_visible(False)

    ax.set_title(f'{asset_data_file.removesuffix(".xlsx")}')

    if save:
        return save_result(ax, f'Asset_Data_PLOT', save_directory)
    else:
        return ax

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Reads constraint resolution for a specific asset, specific scenario, and specific batch across each version number.'''
def read_conres_by_batch(PSARunID, directory, batch='latest', save=False):
    directory = default_directory(PSARunID, directory=directory)
    fldr = f'{directory}\\{PSARunID}_PSA_CANDIDATE-RESPONSES.xlsx'
    file_list = fldr.split('.')
    df_old = []

    file = file_list[0]
    if type(batch) == str:
        batch = latest_batch(PSARunID, directory)

    file = f'{file}-V{batch}-0'

    extension = file_list[1]
    i = int(file[-1])

    if not os.path.isfile(f'{file}.{extension}'):
        raise FileNotFoundError(f'File does not exist: {file}.{extension}')
    while os.path.isfile(f'{file}.{extension}'):
        df_new = pd.read_excel(f'{file}.{extension}', usecols=[0,6,9,10,12,13,14])
        df_new['Version'] = i
        try:
            df_old = pd.concat([df_old,df_new])
        except:
            df_old = df_new.copy()
        finally:
            i += 1
            file = f'{file[:-1]}{i}'

    df_old = df_old.rename(columns={'req_id':'Req_ID'})
    df_old = df_old.set_index(['Version','Req_ID']).sort_index()
    return df_old

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Plots constraint resolution for a specific asset, specific scenario, and specific batch across each version number.'''
def plot_conres_by_batch(PSARunID, directory, batch='latest', scenario='BASE', asset=None, plot=True, save=False):
    directory = default_directory(PSARunID, directory=directory)
    if type(batch) == str:
        batch = get_sf_dict(PSARunID, directory)[scenario][0]
    df_old = read_conres_by_batch(PSARunID, directory, batch)
    
    if plot:
        c_dict = {}
        si = get_initials(scenario)

        if asset == None:
            asset = df_old.loc[0,'constrained_pf_id'].iloc[0] # Default to first instance of an asset

        df_old = df_old[(df_old['constrained_pf_id']==asset) & (df_old['scenario']==scenario)]
        power_df = df_old['residual_req_power_kw'].unstack(level=0)

        for version in list(power_df):
            alpha = 1.0/(len(list(power_df)) - int(version)) # Assumes we start from version 0 to avoid ZeroDivision error
            c_dict[version] = (0.1,0.2,0.3,alpha) # RGBA tuple

        fig, ax = plt.subplots(nrows=1, ncols=2)

        ax = power_df.plot(kind='bar', rot=0, color=c_dict)
        ax.set_title(f'{PSARunID}_PSA_CANDIDATE-RESPONSES-V{batch}-n\n{scenario}, {asset}')
        ax.set_ylabel('Required kW')
        ax.grid(axis='y')

        if save:
            return save_result(fig, f'{si}_{asset}_{batch}_by_Version_GRAPH', directory)
        else:
            return ax

    if save:
        return save_result(df_old, f'{si}_{asset}_{batch}_ConRes_by_Version_DATA', directory)
    else:
        return df_old

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Validates that a certain column in a dataframe A: exists and B: contains only the listed values'''
def plot_multiple_conres_by_batch(PSARunID, directory, asset=None, scenario='BASE', batch='latest', save=False):
    directory = default_directory(PSARunID, directory=directory)
    if type(batch) == str:
        batch = get_sf_dict(PSARunID, directory)[scenario][0]

    df_old = read_conres_by_batch(PSARunID, directory, batch)
    if not asset:
        asset = df_old.loc[0,'constrained_pf_id'].iloc[0] # Default to first instance of an asset
    df_old = df_old[(df_old['constrained_pf_id']==asset) & (df_old['scenario']==scenario)]

    if df_old.shape[0] == 0:
        raise FileNotFoundError(f'Either asset: {asset} or scenario: {scenario} not present in ConRes files')
    
    power_df = df_old['residual_req_power_kw'].unstack(level=0)
    load_df = df_old['loading_pct'].unstack(level=0)

    x = load_df.index
    width = 0.75/len(list(load_df))
    c_dict = {}
    multiplier = 0
    si = get_initials(scenario)
    threshold = get_asset_threshold(asset, PSARunID, directory)

    for version in list(load_df):
        alpha = 1.0/(len(list(load_df)) - int(version)) # Assumes we start from version 0 to avoid ZeroDivision error
        c_dict[version] = (0.1,0.2,0.3,alpha) 

    fig, ax = plt.subplots(nrows=1, ncols=2, figsize=(10,6))#nrows=1, ncols=2
    fig.suptitle(f'{PSARunID}_PSA_CANDIDATE-RESPONSES-V{batch}-n\n{scenario}, {asset}')    

    for version in list(load_df):
        offset = width * multiplier
        ax[1].bar(x + offset, load_df[version], width, color=c_dict[version], label=f'V{version}')
        ax[0].bar(x + offset, power_df[version], width, color=c_dict[version], label=f'V{version}')
        multiplier += 1

    ax[1].axhline(y=threshold, color='red')

    ax[0].set_xticks(x + offset/2, load_df.index)
    ax[1].set_xticks(x + offset/2, load_df.index)

    ax[0].set_ylabel('Required kW')
    ax[1].set_ylabel('% Loading')

    ax[0].set_xlabel('Req_ID')
    ax[1].set_xlabel('Req_ID')

    ax[0].set_title('Required kW')
    ax[1].set_title('% Loading')

    ax[0].grid(axis='y')
    ax[1].grid(axis='y')

    ax[1].legend(bbox_to_anchor=(1.025,1),loc='upper left')
    
    if save:
        return save_result(fig, f'{si}_{asset}_{batch}_by_Version_Multiple_GRAPH', directory)
    else:
        return fig
    
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Validates that a certain column in a dataframe A: exists and B: contains only the listed values'''
def validate_df(df, column, contains=None):
    try:
        df_list = df[column].drop_duplicates().to_list()
    except:
        raise KeyError(f'{column} not in DataFrame. Columns available: {list(df)}')
    if bool(contains):
        broken_values = []
        contains = [contains] if type(contains) != list else contains
        for value in df_list:
            if value not in contains:
                broken_values.append(value)
    
    if bool(broken_values):
        raise ValueError(f'Unexpected values in {column}: {(", ").join(broken_values)}\n\nExpected values: {(", ").join(contains)}')
    return True

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Detailed Validation of asset and asset_type columns of a dataframe such that the index is asset_names and there is a column asset_type'''
def detailed_validate_assets(df, ref_dict={}, file='File'):
    if not isinstance(df, pd.DataFrame):
        raise TypeError('Invalid argument for df in pa.detailed_validate_assets()')
    if not isinstance(ref_dict, dict):
        raise TypeError('Invalid argument for ref_dict in pa.detailed_validate_assets()')
    
    columns = list(df)

    if df.index.name == None or df.index.name not in ['SIA_ID','NERDA_ID','ASSET_ID']:
        id_column = [id_column for id_column in ['SIA_ID','NERDA_ID','ASSET_ID'] if id_column in columns]
        if bool(id_column):
            df = df.set_index(id_column[0])            
        else:
            raise TypeError(f'Column ASSET_ID not in {file}. Available Columns:\n{(", ").join(columns)}')
    if 'ASSET_TYPE' not in columns:
        raise TypeError(f'Column ASSET_TYPE not in {file}. Available Columns:\n{(", ").join(columns)}')
    
    if not bool(ref_dict):
        dict_config = ut.PSA_SND_read_config(const.strConfigFile)
        RefAssetFile = f'{dict_config[const.strWorkingFolder]}/{const.strRefAssetFile}'
        ref_dict = dict(zip([const.strAssetLine,const.strAssetTrans,const.strAssetGen,const.strAssetSwitch,const.strAssetCoupler],\
                            fv.readRefAssetNames(RefAssetFile)))
    
    broken_pf_dict, broken_type_dict = {}, {}
    for asset, asset_type in zip(list(df.index),list(df['ASSET_TYPE'])):
        try:
            if asset not in ref_dict[asset_type]:
                try:
                    broken_pf_dict[asset_type].append(asset)
                except:
                    broken_pf_dict[asset_type] = [asset]
        except:
            try:
                broken_type_dict[asset_type].append(asset)
            except:
                broken_type_dict[asset_type] = [asset]

    msg = ''      
    if bool(broken_pf_dict):
        msg += f"The following assets in {file} were not found in PowerFactory model."
        for key, value in broken_pf_dict.items():
            msg += f'\n{key}: {(", ").join(value)}'
    if bool(broken_type_dict):  
        msg += f"\n\nThe following assets in {file} have an invalid asset_type."
        for key, value in broken_type_dict.items():
            msg += f'\n{key}: {(", ").join(value)}'
    if bool(msg):
        raise ValueError(msg)
    
    return True

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Analysis of PSA SND Interaction, this function collects raw data from a results folder'''
def get_file_creation_times(directory, stopping_index=30, target=['AUT','MAN'], save=False):
    filtered_dir = [identity for identity in os.listdir(directory) if identity[0:3] in target and len(identity)==20]
    sorted_dir = sorted(filtered_dir, key = lambda t: datetime.strptime(t[4:], '%Y-%m-%d-%H-%M'), reverse=True)
    #print(sorted_dir[:stopping_index])
    df = pd.DataFrame(columns=['psarunid','file','date_created','date_modified','size_kb'])
    for PSARunID in sorted_dir[:stopping_index]:
        fldr = f'{directory}\\{PSARunID}'
        files = [file for file in os.listdir(fldr) if file[-5:] == '.xlsx']
        if len(files) > 2: # Only accept folders containing constraint resolution
            for file in files:
                file_path = f'{fldr}\\{file}'
                ct_obj = time.strptime(time.ctime(os.path.getctime(file_path)))
                creation_time = time.strftime("%Y-%m-%d %H:%M:%S", ct_obj)
                mt_obj = time.strptime(time.ctime(os.path.getmtime(file_path)))
                modified_time = time.strftime("%Y-%m-%d %H:%M:%S", mt_obj)
                file_size = os.path.getsize(file_path) / 1000
                df.loc[len(df)] = [PSARunID,file,creation_time,modified_time, file_size]
                
    df.set_index(['psarunid','file'], inplace=True)
    if save:
        return save_result(df, 'PSA_SND_File_Timings', directory, new_folder=False)
    else:
        return df
    
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Analysis of PSA SND Interaction, this function generates various excel files based on above input'''
def file_creation_time_delta(df, PSARunID_single=None, save=False, agg=True, file_type='both', directory=None):
    if not directory:
        dict_config = ut.PSA_SND_read_config(const.strConfigFile)
        directory = f'{dict_config[const.strWorkingFolder]}{const.strDataResults}'

    df = df.reset_index().sort_values(by='date_created')
    df['psarunid'] = [('_').join(file.split('_')[0:2]) for file in list(df['file'])]
    column_name = 'Minutes_Taken (PSA/SND)'
    final_dir = 'PSA_SND_file_timing_analysis'
    
    if PSARunID_single == None:
        values = list(df['psarunid'].unique())
    elif PSARunID_single in list(df['psarunid'].unique()):
        values = [PSARunID_single]
    else:
        raise ValueError(f'PSARunID: {PSARunID_single} not found in reference dataframe')
        
    if file_type == 'both':
        v, length = True, 0 
    elif file_type == 'SF':
        v, length = False, 2 #V0 is 2 characters, a trait of all SF files
        column_name += ' SF only'
        final_dir += '_SF'
    elif file_type == 'CR':
        v, length = False, 4 #V0-1 is 4 characters, a trait of all CR files
        column_name += ' CR only'
        final_dir += '_CR'
    else:
        raise ValueError('file_type can only be literal: both, SF, CR')
    
    df_final = pd.DataFrame(columns=['PSARunID','File_Version',column_name])
    for PSARunID in values:
        df_filt = df[df['psarunid'] == PSARunID].reset_index(drop=True)
        if len(df_filt)//2 != len(df_filt)/2:
            df_filt.drop(len(df_filt)-1, inplace=True)
            #print(f'Unmatched File in {PSARunID}')
        for i in range(2, len(df_filt), 2): # Start at 2 to skip constraints and flex-reqts files that will always be made first
            version = f"V{df_filt.loc[i,'file'].split('V')[1].removesuffix('.xlsx')}"
            if v or len(version) == length:
                time_delta = df_filt.loc[i+1,'date_created'] - df_filt.loc[i,'date_created']
                file_size = (df_filt.loc[i+1,'size_kb'] - df_filt.loc[i,'size_kb']) / 2
                minutes_diff = time_delta.seconds / 60
                df_final.loc[len(df_final)] = [PSARunID, version, minutes_diff]

    if agg:
        df_final = df_final.drop('File_Version',axis=1).groupby('PSARunID').agg(['sum','min','mean','max','std','count'])
        df_final = df_final.rename(columns={'count':'file_count'})
        final_dir += '_agg'
    if save:
        return save_result(df_final, final_dir, directory, new_folder=False)
    else:
        return df_final

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Plots the loading of all the assets within one PSARunID on one graph'''
def read_asset_overview(PSARunID, directory, scenario, df=False, save=False):
    directory = default_directory(PSARunID, directory)
    if isinstance(df, bool):
        df_overview = overview(PSARunID, scenario, directory).data
    elif isinstance(df, pd.DataFrame):
        df_overview = df
    else:
        df_overview = df.data

    df_rq = []
    df_ld = []
    
    for value in df_overview.index:
        asset = df_overview.loc[value,'Asset Name']
        asset_type = df_overview.loc[value,'Type']
        try:
            df_temp_rq = read_flex_rq(asset, PSARunID, directory, scenario, plot=False)
            df_temp_ld = read_asset(asset, asset_type, PSARunID, scenario, directory, plot=False, decorate=False)
        except FileNotFoundError as err:
            print(err)
            continue
        
        for i, df in ((0, df_temp_rq),(1, df_temp_ld)):
            df_melt = pd.melt(df, var_name='Day', value_name=asset)
            df_melt['HH'] = df_melt.index % 48
            df_melt = df_melt.set_index(['Day','HH'])
            if i == 0:
                try:
                    df_rq = df_melt.join(df_rq, ['Day','HH'], how='left')
                except BaseException:
                    df_rq = df_melt.copy()
            elif i == 1: 
                try:
                    df_ld = df_melt.join(df_ld, ['Day','HH'], how='left')
                except BaseException:
                    df_ld = df_melt.copy()
            
    df_rq = df_rq.reset_index().drop(columns='HH')
    df_ld = df_ld.reset_index().drop(columns='HH')

    fig, ax = plt.subplots(2,1, sharex=True, figsize=(10,5))
    fig.suptitle(f'Assets In {PSARunID}')
    
    plot_generic_graph(PSARunID, df_rq, y_label='Req_kW', title='', decorate='type_2', vary='colour', axes=ax[0])
    legend_elements = plot_generic_graph(PSARunID, df_ld, y_label='% Loading', title='', decorate='type_2', vary='colour', axes=ax[1])

    fig.legend(bbox_to_anchor=(1, 0.88), loc='upper center', handles=legend_elements, borderaxespad=0)

    if save:
        return save_result(ax, f'{scenario}_all_asset_loading', directory)
    else:
        return ax
    
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Takes a start, stop, and step and returns all the runs that fall within that start stop and step'''
def get_applicable_runs(start='earliest', stop='latest', step=1, directory=None, target='AUT'):
    if not directory:
        dict_config = ut.PSA_SND_read_config(const.strConfigFile)
        directory = f'{dict_config[const.strWorkingFolder]}{const.strDataResults}'
    
    if target == 'both':
        target = ['AUT','MAN']
    else:
        target = [target]
    
    if type(start) == list:
        sorted_dir = sorted(start, key = lambda t: datetime.strptime(t[4:], '%Y-%m-%d-%H-%M'), reverse=False)
    else:
        filtered_dir = [identity for identity in os.listdir(directory) if identity[0:3] in target and len(identity)==20]
        sorted_dir = sorted(filtered_dir, key = lambda t: datetime.strptime(t[4:], '%Y-%m-%d-%H-%M'), reverse=False)
        if not bool(sorted_dir):
            raise KeyError('No Applicable PSARunIDs')
        dict_dir = dict(zip(sorted_dir, list(range(0,len(sorted_dir)))))
        #print(dict_dir)
        if stop != 'latest':
            end = dict_dir[stop] + 1
            #print(end)
        else:
            end = len(sorted_dir)
        if start != 'earliest':
            begin = dict_dir[start]
            #print(begin)
        else:
            begin = 0
        sorted_dir.reverse()
        sorted_dir = sorted_dir[begin:end:step]
           
    return sorted_dir

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Limits a listdir of data/results'''
def limit_runs(runs, start, limit=11, daily_flag=False):
    applicable_runs = []
    i = 0
    fldr = runs[i]
    dt_object = datetime.strptime(start[4:], '%Y-%m-%d-%H-%M')
    diff = (dt_object - datetime.strptime(fldr[4:], '%Y-%m-%d-%H-%M')).days
    old_diff = -1
    while diff < limit:
        if diff >= 0 and not daily_flag:
            applicable_runs.append(fldr)
        elif diff >= 0 and diff > old_diff and daily_flag:
            applicable_runs.append(fldr)
        i += 1
        if i > len(runs)-1:
            break
        else:
            fldr = runs[i]
            old_diff = diff
            diff = (dt_object - datetime.strptime(fldr[4:], '%Y-%m-%d-%H-%M')).days
    
    return applicable_runs

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Plots a single asset across multiple PSARunIDS'''
def plot_inter_run_flex_rq(directory, asset=None, start='earliest', stop='latest', step=1, scenario='BASE', save=False, title=None, decorate='type_2'):
    if type(start) == list:
        sorted_dir = start
    else:
        sorted_dir = get_applicable_runs(start, stop, step, directory)

    r = 2
    first = sorted_dir[0]
    if len(sorted_dir) > 1:
        runs = limit_runs(sorted_dir, first)
    else:
        runs = sorted_dir
    
    if asset == None:
        scenario_dict = {'BASE': const.strBfolder,
                     'MAINT': const.strMfolder,
                     'CONT': const.strCfolder,
                     'MAINT_CONT': const.strMCfolder}
        for folder in ['LN_LDG_DATA','TX_LDG_DATA']:
            try:
                asset = os.listdir(f'{directory}\\{first}{scenario_dict[scenario]}{folder}')[0].split('-')[1].removesuffix('.xlsx')
                asset_type = const.strAssetTrans if folder == 'TX_LDG_DATA' else const.strAssetLine
            except BaseException as err:
                r += 1
                continue
    else:
        asset_type = get_asset_type(first, directory, asset)
                   
    if asset == None:
        raise ValueError('Could not default to asset and no asset specified')

    df_rq = []
    df_ld = []
    for run in runs:
        try:
            df_temp_rq = read_flex_rq(asset, run, f'{directory}\\{run}', scenario, plot=False)
            df_temp_ld = read_asset(asset, asset_type, run, scenario, f'{directory}\\{run}', plot=False, decorate=False)
        except FileNotFoundError as err:
            print(err)
            continue
        
        for i, df in ((0, df_temp_rq),(1, df_temp_ld)):
            value_name = asset if len(runs) == 1 else run
            df_melt = pd.melt(df, var_name='Day', value_name=value_name)
            df_melt['HH'] = df_melt.index % 48
            df_melt = df_melt.set_index(['Day','HH'])
            if i == 0:
                try:
                    df_rq = df_melt.join(df_rq, ['Day','HH'], how='left')
                except BaseException:
                    df_rq = df_melt.copy()
            elif i == 1: 
                try:
                    df_ld = df_melt.join(df_ld, ['Day','HH'], how='left')
                except BaseException:
                    df_ld = df_melt.copy()

    if isinstance(df_rq, list) or isinstance(df_ld, list):
        raise FileExistsError(f'None of the runs in {(", ").join(runs)} produced results for asset {asset} in scenario {scenario}.\
                              \nCheck for RQ: {not isinstance(df_rq, list)}\
                              \nCheck for LD: {not isinstance(df_ld, list)}')
    
    df_rq = df_rq.reset_index().drop(columns='HH')
    df_ld = df_ld.reset_index().drop(columns='HH')
    threshold = get_asset_threshold(asset, runs[0], directory)

    fig, ax = plt.subplots(2,1, sharex=True, figsize=(10,5))
    if not title:
        title = f'{asset} Across PSA Runs'
    
    fig.suptitle(title)

    if len(runs) > 1:
        colour = ['purple'] + ['blue']*(len(runs)-r) + ['purple']
    else:
        colour = 'blue'
    
    plot_generic_graph(run, df_rq, y_label='Req_kW', title='', decorate=decorate, colour=colour, vary='alpha', axes=ax[0])
    legend_elements = plot_generic_graph(run, df_ld, y_label='% Loading', title='', decorate=decorate, colour=colour, vary='alpha', axes=ax[1], h_line=threshold)
    
    if len(legend_elements) < 20:
        fig.legend(bbox_to_anchor=(1.12, 0.88), handles=legend_elements, borderaxespad=0)

    if save:
        return save_result(fig, f'{asset}_loading_plot_across_runs', directory, new_folder=True)
    else:
        return ax
    
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Plots SIA data for a single primary across multiple PSARunIDS'''
def plot_inter_run_sia(directory, primary, start='earliest', stop='latest', step=1, save=False):
    if type(start) == list:
        runs = start
        first = runs[0]
    else:
        runs = get_applicable_runs(start, stop, step, directory)
        first = runs[0]
        runs = limit_runs(runs, first)

    r = 2
    df_final = []
    for run in runs:
        try:
            df = plot_sia(primary, run, directory, plot=False)
        except FileNotFoundError:
            r += 1
            continue
            
        df = df.rename(columns={'pct_loading':run})
        df['HH'] = df.index % 48
        df = df.set_index(['Day','HH'])
        try:
            df_final = df.join(df_final, ['Day','HH'], how='left')
        except BaseException:
            df_final = df.copy()

    threshold = get_asset_threshold(primary, runs[0], directory)
    df_final = df_final.reset_index().drop(columns='HH')
    if len(runs) > 1:
        colour = ['purple'] + ['darkred']*(len(runs)-r) + ['purple']
    else:
        colour = 'darkred'

    title = f'{primary} SIA Data Across PSA Runs'
    ax = plot_generic_graph(run, df_final, title=title, y_label='% Loading', decorate='type_2', vary='alpha', colour=colour, h_line=threshold)
        
    if save:
        return save_result(ax, f'{primary}_SIA_plot_across_runs', directory, new_folder=True)
    else:
        return ax
    
# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Plots NeRDA data across multiple PSARunIDS'''
def plot_inter_run_nerda(directory, start='earliest', stop='latest', step=1, save=False):
    if type(start) == list:
        sorted_dir = start
    else:
        sorted_dir = get_applicable_runs(start, stop, step, directory)

    first = sorted_dir[0]
    runs = limit_runs(sorted_dir, first)

    df_final = []
    
    for run in runs:
        file = f'{directory}\\{run}{const.strNeRDAfolder}{run}-NeRDA_DATA.xlsx'
        try:
            df = pd.read_excel(file, usecols=[0,1]).rename(columns={'status':run})
            df = df.drop_duplicates(subset='shortName').set_index('shortName')
        except FileNotFoundError:
            continue
        try:
            df_final = df.join(df_final, 'shortName', how='left')
        except BaseException:
            df_final = df.copy()

    for idx in df_final.index:
        if len(set(df_final.loc[idx])) == 1:
            df_final.drop(index=idx, inplace=True)

    if df_final.shape[0] == 0:
        raise ValueError(f'NeRDA Data Remains Constant in the specified files: {(", ").join(runs)}')
    
    if df_final.shape[0] == 0:
        raise ValueError(f'NeRDA Data Remains Constant in the specified runs {(", ").join(runs)}')

    if save:
        return save_result(df_final, f'NeRDA_Data_Across_Runs', directory, new_folder=True)
    else:
        return df_final

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''Stitches together V0-X, V1-X, V2-X etc where X is in an integer to create a flex-reqts like file'''
def stitch_iterations(PSARunID, directory):
    sf_files = [file[-6] for file in os.listdir(directory) if const.strPSAResponses in file]
    if not sf_files:
        raise FileNotFoundError(f'No SF Files in {directory}\\{PSARunID}')

    file_dict, df_dict = {}, {}
    #debug_dict = {}
    i, x = 0, 0
    while x < len(sf_files):
        x, df = 0, []
        for batch in sf_files:
            file = f'{directory}\\{PSARunID}_PSA_CANDIDATE-RESPONSES-V{batch}-{i}.xlsx'
            if os.path.isfile(file):
                file_dict[batch] = file
            elif batch in list(file_dict.keys()):
                file = file_dict[batch]
                x += 1
            else:
                x += 1
                break
            df_temp = pd.read_excel(file)
            #print(df_temp.shape)
            df_temp.insert(0, 'batch', [batch]*len(df_temp))
            try:
                df = pd.concat([df, df_temp], ignore_index=True)
            except BaseException:
                df = df_temp.copy()
        #print(df.shape)
        if x < len(sf_files):
            #debug_dict[i] = df.groupby(['batch','start_time']).agg('size').to_frame()
            df_dict[i] = df.sort_values(by='req_id', ignore_index=True)
            i += 1

    return df_dict