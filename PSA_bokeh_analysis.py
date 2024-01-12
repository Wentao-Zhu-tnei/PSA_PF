import os
import pandas as pd
import PSA_analysis as pa
#import numpy as np
import PSA_SND_Constants as const
from datetime import datetime, timedelta
from bokeh.plotting import figure
from bokeh.models import ColumnDataSource, HoverTool, GroupFilter, CDSView

# -------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'''This works identically to read_flex_rq in PSA_analysis, but produces an interactive Bokeh plot in Jupyter'''
def bokeh_flex_rq(asset, PSARunID, directory=None, scenario='BASE', plot=True, plot_residual=True):
    style_dict = {'Day ahead': 'dashed',
                  'Week ahead': 'solid',
                  'Weekend': 'dotted',
                  'None': 'dotted'}
    
    directory = pa.default_directory(PSARunID, directory)

    scenario = '_'.join(scenario.split(' ')).upper()
    si = pa.get_initials(scenario)

    flex_reqts = pa.find_data(PSARunID, 'FLEX_REQTS', directory=directory, scenario=scenario)
    super_df = pd.read_excel(os.path.join(flex_reqts, f'{PSARunID}{const.strPSAFlexReqts}.xlsx'))

    date = PSARunID.split('_')[1]
    dt_object = datetime.strptime(date, '%Y-%m-%d-%H-%M')

    super_df['start_time'] = pd.to_datetime(super_df['start_time'])
    super_df['Day'] = super_df['start_time'].apply(pa.delta, start_date=dt_object)
    super_df['Day_HH'] = super_df['start_time'].dt.hour * 2 + super_df['start_time'].dt.minute / 30
    super_df['Day_HH'] = super_df['Day_HH'].astype(int)
    super_df['HH'] = super_df['Day']*48 + super_df['Day_HH']

    asset_df = super_df[super_df['constrained_pf_id'] == asset]
    pivot_df = asset_df[['required_power_kw','Day','Day_HH']].pivot(index='Day_HH',columns='Day', values='required_power_kw')
    if pivot_df.shape[0] != 48:
        missing_values = [value for value in list(range(0, 48)) if value not in list(pivot_df.index.to_series())]
        missing_dict = {}

        for value in missing_values:
            missing_dict[value] = [0]*pivot_df.shape[1]

        filler_df = pd.DataFrame.from_dict(missing_dict, orient='index', columns=list(pivot_df))
        pivot_df = pd.concat([pivot_df, filler_df])
        pivot_df.sort_index(inplace=True)

    pivot_df.set_index(pivot_df.index.map(pa.HH_to_timestamp), inplace=True)
    pivot_df.fillna(0, inplace=True)

    day_max = 11 # max(list(pivot_df))
    missing_columns = [number for number in range(0,day_max) if number not in list(pivot_df)]
    for column in missing_columns:
        pivot_df[column] = 0
    pivot_df = pivot_df.loc[:,list(range(0,day_max))]

    
    if plot:
        column_names = pa.convert_to_dates(PSARunID, pivot_df)

        df_melt = pd.melt(pivot_df, var_name='Day', value_name='Req_kw')
        df_melt['HH'] = df_melt.index
        df_melt.set_index(['Day','HH'], inplace = True)
         
        test_df = asset_df.loc[:,['Day','HH','req_id']]
        test_df.set_index(['Day','HH'], inplace = True)
        
        if plot_residual:
            try:
                df1 = pa.read_conres(PSARunID, directory, save=False).loc[asset]
                # ^ Filter for asset
                df2 = df1.loc[(df1.index.get_level_values('Unnamed: 0') == 'final'),['res_flex_req_power_kw','tot_offered_power_kw']]
                # ^ Filter only final power
                df2 = df2.reindex(df2['res_flex_req_power_kw'].abs().sort_values().index) 
                # ^ sort by smallest to largest in absolute terms
                df2 = df2.reset_index().drop_duplicates('req_id').set_index('req_id') 
                # ^ drop duplicate values and reindex (for dealing with multiple res_ids to one req_id)
                test_df = test_df.merge(df2, how='left', left_on='req_id', right_index=True)
            except FileNotFoundError:
                plot_residual = False
            except BaseException as err:
                print(f'Plot_resudual Error in plot_bokeh_analysis: {err}')
                plot_residual = False
                
        df_merged = df_melt.merge(test_df, how='left', left_index=True, right_index=True)
        df_merged['HH'] = df_merged.index.levels[1]
        df_merged['HH_debug'] = list(range(0,48))*day_max
        df_merged['HH_debug'] = df_merged['HH_debug'].apply(pa.HH_to_timestamp)
    
        # Graph Plotting
        ax = figure(width=900, height=400)
        data = ColumnDataSource(df_merged)
        
        labels = list(column_names.values())
        end_date = list(column_names.values())[-1].split(' ')[1]
        dt_object = datetime.strptime(end_date, '%d/%m/%y') + timedelta(days = 1)
        labels.append(dt_object.strftime('%A')[:3] + ' ' + dt_object.strftime('%d/%m/%y'))

        date = PSARunID.split('_')[1]
        dt_object = datetime.strptime(date, '%Y-%m-%d-%H-%M')
        xticks = [0]

        x0 = df_melt.index.levels[1]
        y0 = df_melt['Req_kw']

        y_max = max(y0) * 1.1
        y_min = min(y0) * 1.1

        plot1 = ax.line(source=data, x='HH',y='Req_kw', color='white')
        ax.add_tools(HoverTool(renderers=[plot1], tooltips=[('kW','@Req_kw'),('HH','@HH_debug'),('Req_ID','@req_id')], mode='vline'))
        
        if plot_residual:
            plot2 = ax.line(source=data, x='HH',y='res_flex_req_power_kw', color='darkred')
            ax.add_tools(HoverTool(renderers=[plot2], tooltips=[('res_kW','@res_flex_req_power_kw'),('resp_id','@resp_id'),('off_pwr','@tot_offered_power_kw')], mode='vline'))

        for i in range(len(list(column_names.values()))):
            xticks.append(48 * (i + 1))
            df_date = df_melt.loc[i]
            x = df_date.index
            y = df_date['Req_kw']
            auction = pa.auction_relevance(i, dt_object)
            auction = 'None' if auction == 'Weekend' else auction
            style = style_dict[auction]
            ax.line(x=x, y=y, line_dash=style, color='blue', legend_label=auction)

            ax.multi_polygons(xs=[
                                    [[ [30+48*i,30+48*i,38+48*i,38+48*i] ]], # SPM Window 
                                    [[ [20+48*i,20+48*i,28+48*i,28+48*i] ]]  # SEPM Window
                                 ],
                              ys=[
                                    [[ [y_min,y_max,y_max,y_min] ]], # SPM Window
                                    [[ [y_min,y_max,y_max,y_min] ]]  # SEPM Window
                                 ],
                              color = ['green','yellow'],
                              alpha = [0.2,0.2])

        final_dict = dict(zip(xticks, labels))    

        ax.add_tools(HoverTool(renderers=[plot1], tooltips=[('kW','@Req_kw'),('HH','@HH_debug'),('Req_ID','@req_id')], mode='vline'))
        ax.xaxis.ticker = xticks
        ax.xaxis.major_label_overrides = final_dict
        ax.xaxis.major_label_orientation = 45
        ax.title.text = f'{PSARunID} | {asset}'
        ax.title.text_font_size = '14pt'
        ax.yaxis.axis_label = 'Required kW'
        ax.yaxis.axis_label_text_font_style = 'bold'
        
        return ax
    else:
        return pivot_df