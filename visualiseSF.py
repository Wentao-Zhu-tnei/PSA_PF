import os
import glob
import pandas as pd
import matplotlib.pyplot as plt

def find_aut_folders(target_folder):
    return [os.path.join(target_folder, f) for f in os.listdir(target_folder) if os.path.isdir(os.path.join(target_folder, f)) and f.startswith('AUT')]


def find_psa_workbooks(folder):
    return glob.glob(os.path.join(folder, '*_PSA_SF-RESPONSES*.xlsx'))


def read_first_sheet(workbook):
    return pd.read_excel(workbook, sheet_name=0)


def filter_columns(df):
    return df[['constrained_pf_id', 'flex_pf_id', 'sensitivity_factor', 'sensitivity_factor_dir', 'scenario']]


def add_run_id_column(df, run_id):
    df['PSArunID'] = run_id
    return df

def add_directed_sensitivity_factor_column(df):
    df = df[(abs(df['sensitivity_factor']) <= 1) & (abs(df['sensitivity_factor_dir']) <= 1)]
    df_mod = df.copy()
    df_mod['directed_sensitivity_factor'] = df['sensitivity_factor'] * df['sensitivity_factor_dir']
    df_mod = df_mod.drop(columns=['sensitivity_factor', 'sensitivity_factor_dir'])
    return df_mod

def sort_dataframe_by_constrained_pf_id(df):
    return df.sort_values(['constrained_pf_id', 'flex_pf_id', 'PSArunID'])

def aggregateSF(target_folder):
    dfs = []
    for folder in find_aut_folders(target_folder):
        for workbook in find_psa_workbooks(folder):
            df = read_first_sheet(workbook)
            df = filter_columns(df)
            df = add_run_id_column(df, os.path.basename(folder))
            df = add_directed_sensitivity_factor_column(df)
            dfs.append(df)
    df = pd.concat(dfs, ignore_index=True)
    df = sort_dataframe_by_constrained_pf_id(df)
    return df

def plot_constrained_flex_relationship(df):

    cmap = plt.cm.RdBu_r

    plt.figure(figsize=(10, 8))
    plt.scatter(x='constrained_pf_id', y='flex_pf_id', c='directed_sensitivity_factor', cmap=cmap, data=df, alpha=0.8)

    cbar = plt.colorbar()
    cbar.set_label('Directed Sensitivity Factor', fontsize=14)

    for x, y, z in zip(df['constrained_pf_id'], df['flex_pf_id'], df['PSArunID']):
        plt.text(x, y, str(z).replace('AUT_',''), fontsize=2)

    plt.title('Constrained-Flex Relationship', fontsize=14)
    plt.ylabel('Flex PF ID', fontsize=12, labelpad=10)
    plt.xlabel('Constrained\nPF ID', fontsize=12, labelpad=10, rotation=0, va='center')

    plt.xticks(fontsize=8, rotation=90)
    plt.yticks(fontsize=8)

    plt.show()

def calculate_stats(df):
    # calculate the min, max, and mean values of the 'directed_sensitivity_factor' column for each group
    stats = df.groupby(['constrained_pf_id', 'flex_pf_id','scenario']).agg({'directed_sensitivity_factor': ['min', 'max', 'mean', 'std']}).reset_index()
    stats.columns = stats.columns.map('_'.join)
    for i in range(3):
        stats.rename(columns={list(stats.columns)[i]: list(stats.columns)[i].rstrip('_')}, inplace=True)

    return stats


def group_and_count(df):
    # return df.groupby(['constrained_pf_id', 'flex_pf_id', 'scenario'],as_index=False).size().reset_index(name='count')
    return df.groupby(['constrained_pf_id', 'flex_pf_id', 'scenario'],as_index=False).size()

def write_to_excel(df, result, target_file):
    # writer = pd.ExcelWriter(target_file, engine='openpyxl')
    writer = pd.ExcelWriter(target_file, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    result.to_excel(writer, sheet_name='Sheet2', index=False)
    writer.close()

def analyseData(target_file):
    df = pd.read_excel(target_file, sheet_name=0)
    stats = calculate_stats(df)
    grouped = group_and_count(df)
    result = pd.merge(grouped, stats, on=['constrained_pf_id', 'flex_pf_id', 'scenario'], how='left')
    write_to_excel(df, result, target_file)

def fullProcess(target_folder, target_file):
    df = aggregateSF(target_folder)
    df.to_excel(os.path.join(target_folder, target_file),index=False)
    # plot_constrained_flex_relationship(df)
    analyseData(os.path.join(target_folder, target_file))

target_folder = r"C:\Wentao Zhu\Project\PSA\Networks-Transition-Technical-Trials-PSA-1\data\results"
target_file = r"SFs.xlsx"
fullProcess(target_folder,target_file)