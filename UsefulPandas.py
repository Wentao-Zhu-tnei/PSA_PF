import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib as mpl
import UsefulPrinting


def filter_df_single(df, col_by, string, regex):
    '''This function looks for a specific word in a specific column of a pandas dataframe'''

    ### Look for all cases where the word is contained
    if regex == True:
        df_filtered = df[df[col_by].str.contains(string + '.*', na=False, case=False, regex=True)]

    ### Look for exact match of the word
    elif regex == False:
        base = r'^{}'  ## {word}
        expr = '(?:\s|^){}(?:,\s|$)'  # expr{}expr
        df_filtered = df[df[col_by].str.contains(base.format(''.join(expr.format(string))), na=False, case=False, regex=True)]

    elif regex is None:
        df_filtered = df[df[col_by].str.contains(string, na=False, case=False,regex=False)]

    return df_filtered


def calc_kVA(kw, kvar):
    kva = (kw ** 2 + kvar ** 2) ** 0.5
    return kva


def df_calc_kVA(df, column_name, kw_col, kvar_col):
    new_col = column_name

    df[new_col] = df.apply(lambda x: calc_kVA(x[kw_col], x[kvar_col]), axis=1)

    #UsefulPrinting.print_spacer()
    #print(f'New column added for kVA: {new_col}')
    #UsefulPrinting.print_spacer()
    return df


def running_diff_abs(arr, ref_start, N):
    '''Return the rolling difference with respect to the first value i.e. absolute'''
    return np.array([arr[i - N] - arr[ref_start] for i in range(N, len(arr) + 1)])


def pd_npsign_alt(num):
    if abs(num) > 0:
        return np.sign(num)
    else:
        return np.nan


def S_LV_diff_sign(df, S_LV, P_LV):
    P_LV_sign = np.sign(df[P_LV])
    df['S_LV_diff_abs'] = running_diff_abs(np.array(df[S_LV]), 0, 1) * P_LV_sign

    return df


def get_direction_flow(P_LV, P_HV):
    '''This function gives you the direction of the flow in a network element'''
    if (abs(P_HV) > abs(P_LV) and np.sign(P_HV - P_LV) > 0):
        sign = np.sign(P_HV - P_LV)
        direction = 'Import'
    elif (abs(P_HV) < abs(P_LV) and np.sign(P_HV - P_LV) < 0):
        sign = np.sign(P_HV - P_LV)
        direction = 'Export'

    return sign, direction


def P_LV_diff_sign(P_LV_diff_abs, P_LV):
    '''This function gives you the direction of the flow between
    the difference in steps and the flow in P_LV_s0 at step 0 (this is always the reference)'''
    if P_LV_diff_abs == 0:
        pf_opposite = np.nan
    elif np.sign(P_LV_diff_abs) == np.sign(P_LV):
        pf_opposite = False
    else:
        pf_opposite = True

    return pf_opposite


def SF_validation(SF, SF_sign, PF_direction_opp, P_LV_s0, Pflex, S_LV):
    ## If the direction of the powerflow at the LV terminal of the network element at the current step (np.sign(Ps0) == np.sign(Psx-Ps0) is the same
    ## as the direction of the flow of the initial step np.sign(Ps0)
    if (PF_direction_opp == True and np.sign(P_LV_s0) == -1):
        # if (PF_direction_opp==False):
        SF_val = abs(SF * SF_sign * Pflex + S_LV)
    else:
        SF_val = abs(SF * SF_sign * Pflex - S_LV)

    return SF_val


def delta_MVA_delta_MW(mva, mw):
    if abs(mva) > 0 and abs(mw) > 0:
        return abs(mva) / abs(mw)
    else:
        return np.nan


def count_occ(df,col):
    '''This function counts the occurrence of a categorical variables in a specific column of a pandas dataframe'''
    df_count=df[col].value_counts(dropna=True)
    df_count_nans=df[col].value_counts(dropna=False)
    print(f'Counting the occurrence of a categorical variables in column [{col}]')
    print('Are there any NaNs?')
    if 'NaN' in df_count_nans.index:
        print(f"Yes! {df_count_nans.loc['NaN']}")
    else:
        print('No!')
        print(f'Number of original unique values (categoricals): {df_count_nans.shape[0]}')
    df_count=df_count.to_frame()
    df_count.loc['total']=df_count.sum(axis=0)
    df_count['proportions']=df_count[col].apply(lambda x: x*100/df_count.loc['total'])
    df_count_max=df_count['proportions'][:-1].max()
    df_count_max_idx=df_count['proportions'][:-1].idxmax()
    print(f'Largest category is [{df_count_max_idx}], with a proportion: {df_count_max:.2f}%')
    return df_count


def get_perc_change(current, previous):
    if current == previous:
        return 0
    try:
        return (abs(current - previous) / previous) * 100.0
    except ZeroDivisionError:
        return float('inf')


def highlight_axis(df, target, axis=1, colour='yellow'):
    '''This function looks for a target (either index name or column name) in the chosen axis and highlights it {colour}'''
    if type(target) != list:
        target = [target]

    if type(df) == pd.DataFrame:
        df = df.style

    # You'll have to imagine rows is columns if you set axis=0, variable names are hard
    return df.apply(lambda rows: [f'background: {colour}' if rows.name in target else '' for row in rows], axis=axis)