import os
import re
import imp
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import UsefulPandas
import powerfactory_interface as pf_interface
import PSA_SND_Utilities as ut
import PSA_SND_Constants as const
import PSA_SIA_Data_Input_Output as sio
import PSA_NeRDA_Data_Input_Output as nrd
from datetime import datetime, timedelta
import math
import PF_Config as pfconf
import PSA_PF_Functions as pffun
import sys
import pathlib
from os.path import exists
import PSA_PF_SF as sf
import PSA_Constraint_Resolution as cnrs
import PSA_Calc_Flex_Reqts as calcflex

psa_run_date='AUT_2023-01-03-23-41'
SND_file=f'{psa_run_date}/{psa_run_date}_SND_RESPONSES.xlsx'
PSA_const_file=f'{psa_run_date}/{psa_run_date}_PSA_CONSTRAINTS.xlsx'

PSAfolder = r"C:\Wentao Zhu\Project\PSA\Networks-Transition-Technical-Trials-PSA-1\data\results\AUT_2023-01-03-23-41"
input_folder = PSAfolder + "\\" + "0 - INPUT_DATA"
NeRDAfolder = PSAfolder + "\\" + "2 - NeRDA_DATA"
PSArunID = "AUT_2023-01-03-23-41"
pf_proj_dir = input_folder
runtime_param_file = PSAfolder + "\\" + PSArunID + const.strRunTimeParams
if exists(runtime_param_file):
    ### Load runtime parameters ###
    dict_params = ut.readPSAparams(runtime_param_file)
    print("Loaded PSArunID parameters")
    print(dict_params)
else:
    status = const.PSAfileExistError
    msg = "Run time Params file doesn't exist: " + PSAfolder + PSArunID + const.strRunTimeParams
dict_config = ut.PSA_SND_read_config(const.strConfigFile)
root_working_folder = dict_config[const.strWorkingFolder]
pf_object = pf_interface.run(dict_config[const.strPFUser])
pf_interface.activate_project(pf_object, dict_params[const.strBSPModel], pf_proj_dir)
dfElmCoup_ALL = pf_object.get_all_circuit_breakers_coup()
dfStaSwitch_ALL = pf_object.get_all_circuit_breakers_switch()
dfStaSwitch_ALL.to_excel(r"C:\Wentao Zhu\Project\PSA\Networks-Transition-Technical-Trials-PSA-1\data\Check\dfStaSwitch_ALL.xlsx")
dfElmCoup_ALL.to_excel(r"C:\Wentao Zhu\Project\PSA\Networks-Transition-Technical-Trials-PSA-1\data\Check\dfElmCoup_ALL.xlsx")
dfNeRDA = pd.read_excel(NeRDAfolder + "\\" + PSArunID +'-NeRDA_DATA.xlsx')
dfNeRDAData = pd.read_excel(r"C:\Wentao Zhu\Project\PSA\Networks-Transition-Technical-Trials-PSA-1\data\results\AUT_2023-01-03-23-41\0 - INPUT_DATA\03-01-23 Asset Data Test JPO SF.xlsx",sheet_name="NeRDA")
dfNeRDACopy = dfNeRDA.copy()
dfNeRDACopy = calcflex.PSA_filter_NeRDA(dfNeRDACopy, dfNeRDAData)
calcflex.PSA_update_switch(pf_object, dfElmCoup_ALL, dfStaSwitch_ALL, dfNeRDACopy)



cwd=os.getcwd()
psa_run_date='AUT_2023-01-03-23-41'
SND_file=f'{psa_run_date}/{psa_run_date}_SND_RESPONSES.xlsx'
PSA_const_file=f'{psa_run_date}/{psa_run_date}_PSA_CONSTRAINTS.xlsx'

verbose=False
bOutput2SND=True
bAuto= False
SND_file=os.path.join('C:/Wentao Zhu/Project/PSA/Networks-Transition-Technical-Trials-PSA-1/data/results/',
                     SND_file)
PSA_const_file=os.path.join('C:/Wentao Zhu/Project/PSA/Networks-Transition-Technical-Trials-PSA-1/data/results/',
                     PSA_const_file)

print('-'*65)
print('Calculating residual flexibility requirements')
print(f'bOutput2SND; [{bOutput2SND}]')
print(f'bAuto; [{bAuto}]')
print(f'SND_file; [{SND_file}]')
print(f'PSA_const_file; [{PSA_const_file}]')
print('-'*65)
status = const.PSAok
dict_config = ut.PSA_SND_read_config(const.strConfigFile)
root_working_folder = dict_config[const.strWorkingFolder]

#get PSArunID from filename
vn, PSArunID, results_folder = sf.getFilePSArunID(SND_file)
PSAfolder = results_folder + "/"
start_date = sf.getPSArunIDdate(PSArunID)
start_time = sf.getPSArunIDtime(PSArunID)
print(f'PSAfolder: [{PSAfolder}]')
print(f'PSAfolder run start_date: [{start_date}]')
print(f'PSAfolder run start_time: [{start_time}]')

if SND_file.find(const.strSNDResponses) >=0:
    bResponse=True
elif SND_file.find(const.strSNDContracts) >=0:
    bResponse = False
else:
    status = const.PSAfileTypeError
    msg = "Unknown file type: " + SND_file

### Load PSArunID folder ###

if exists(PSAfolder + PSArunID + const.strRunTimeParams):
    ### Load runtime parameters ###
    dict_params = ut.readPSAparams(PSAfolder + PSArunID + const.strRunTimeParams)
    print("Loaded PSArunID parameters")
    print(dict_params)
else:
    status = const.PSAfileExistError
    msg = "Run time Params file doesn't exist: " + PSAfolder + PSArunID + const.strRunTimeParams

bDebug = pfconf.bDebug
power_factor = float(dict_params[const.strPowFctr])
lagging = (str(dict_params[const.strPowFctrLag]) == "True")
lDfPSASFResponses = []
file_detector_folder = dict_config[const.strFileDetectorFolder]

#pproject_dir_part = dict_params[const.strBSPModel].replace(dict_params[const.strBSPModel].split("/")[4],"")
#pf_proj_dir = root_working_folder+pproject_dir_part

print("BSP Model = " + dict_params[const.strBSPModel])

### LOAD SND INPUT FILE ###
dfSNDoutput = pd.read_excel(SND_file)
calculate_historical = dfSNDoutput['calculate_historical'][0]
input_folder = PSAfolder + "0 - INPUT_DATA"
pf_proj_dir = input_folder

### LOAD ORIGINAL CONSTRAINTS FILE ###
PSA_const_file = PSAfolder + PSArunID + "_PSA_CONSTRAINTS.xlsx"
dfPSAconstraints = pd.read_excel(PSA_const_file)
dfPSAconstraints['req_id']=dfPSAconstraints['req_id'].astype(str)

### LOAD ASSET DATA FILE ###
inFile = os.path.join(input_folder, dict_params[const.strAssetData])
if exists(inFile):
    dfAssetData = pd.read_excel(os.path.join(input_folder, inFile), sheet_name=const.strAssetSheet)
    dfSIAData = pd.read_excel(os.path.join(input_folder, inFile), sheet_name=const.strSIASheet)
    dfNeRDAData = pd.read_excel(os.path.join(input_folder, inFile), sheet_name=const.strNeRDASheet)
else:
    status = const.PSAfileExistError
    msg = "Asset Data File doesn't exist: " + inFile
    print(status, msg)

NeRDAfolder = PSAfolder+"\\"+"2 - NeRDA_DATA"

if calculate_historical:
    NeRDAfiles = os.listdir(NeRDAfolder)
    for NeRDAfile in NeRDAfiles:
        if NeRDAfile.endswith("NeRDA_DATA.xlsx"):
            NeRDAXls = pd.ExcelFile(NeRDAfolder + "\\" + NeRDAfile)

else:
    status, msg, dfNeRDA = nrd.loadNeRDAData(assets=dfNeRDAData, output_path=NeRDAfolder + "\\" + PSArunID,
                                    user=dict_config[const.strNERDAUsr], pwd=dict_config[const.strNERDAPwd])
    if status != const.PSAok:
        print(status, msg)

pf_object = pf_interface.run(dict_config[const.strPFUser])
pf_interface.activate_project(pf_object, dict_params[const.strBSPModel], pf_proj_dir)

### Get the list of lines and trafos in the network ###
lLines = pf_object.app.GetCalcRelevantObjects("*.ElmLne") ## if this is done somewhere else in the code, there shouldn't be a need to call again
lTrafos_2w = pf_object.app.GetCalcRelevantObjects('*.ElmTr2') ## if this is done somewhere else in the code, there shouldn't be a need to call again
all_loads_df = pf_object.get_loads_names_grid_obj()

pf_object.prepare_loadflow(ldf_mode=dict_config[const.strPFLFMode], algorithm=dict_config[const.strPFAlgrthm],
                           trafo_tap=dict_config[const.strPFTrafoTap],shunt_tap=dict_config[const.strPFShuntTap],
                           zip_load=dict_config[const.strPFZipLoad], q_limits=dict_config[const.strPFQLim],
                           phase_shift=dict_config[const.strPFPhaseShift],
                           trafo_tap_limit=dict_config[const.strPFTrafoTapLim],
                           max_iter=int(dict_config[const.strPFMaxIter]))

print('=' * 65)
ierr1 = pf_object.run_loadflow()
if ierr1 == 0:
    status = const.PSAok
    print(
        f"Load Flow SUCCESSFUL for Initial Conditions\n ")
else:
    status = const.PSAPFError
    print(
        f"Load flow ERROR for Initial Conditions\n ")
print('=' * 65)

all_trafos_2w_df=pf_object.get_transformers_2w_results()
all_trafos_2w_df.head(4)

show = all_trafos_2w_df[['Name', 'Folder', 'Grid','Substation',
                  'From_HV', 'To_LV',
                  'From_HV_brk','To_LV_brk',
                 'From_HV_status', 'To_LV_status', 'CIM_ID' ]].head(20)
print(show)

all_trafos_2w_df[['Name', 'Folder', 'Grid','Substation',
                  'From_HV', 'To_LV',
                  'From_HV_brk','To_LV_brk',
                 'From_HV_status', 'To_LV_status', 'CIM_ID' ]].to_csv(os.path.join(PSAfolder,
                               'all_trafos_2w_df_breakers_new.csv'),
                                       index=False, header=True)

all_lines_df=pf_object.get_lines_results()

show1 = all_lines_df[['Name', 'Grid','From', 'To','From_brk','To_brk', 'From_status', 'To_status', 'CIM_ID' ]]
print(show1)

all_lines_df[['Name', 'Grid','Voltage_level','From', 'To', 'From_brk','To_brk','From_status', 'To_status']].head(20)

all_lines_df[['Name', 'Grid','Voltage_level','From', 'To','From_brk','To_brk','From_status', 'To_status']].\
    to_csv(os.path.join(PSAfolder,'all_lines_df_breakers_new.csv'),index=False, header=True)

pf_object.toggle_switch_ind('COLO_C8L5',action=None,verbose=True)
pf_object.toggle_switch_ind('COLO_C8L5',action='toggle',verbose=True)

pf_object.toggle_switch_ind('COLO_C8L5',action='close',verbose=True)
pf_object.toggle_switch_ind('COLO_C8L5',action='open',verbose=True)

pf_object.toggle_switch_ind('KENN_E1T0',action='open',verbose=True)
pf_object.toggle_switch_ind('BERI_E1T0',action='open',verbose=True)
all_switch_coup=pf_object.get_all_circuit_breakers_coup()
show2 = all_switch_coup.head(5)
print(show2)

pf_object.toggle_switch_coup_ind('UNIS_E2S0',action='open',verbose=True)
pf_object.toggle_switch_coup_ind('COLO_C4S0',action='open',verbose=True)
print(all_switch_coup.shape)

UsefulPandas.filter_df_single(all_switch_coup,'Grid','Cowley Local',regex=True)
print(all_switch_coup['Grid'].unique())

all_switch_coup_11kV_EXL5=UsefulPandas.filter_df_single(all_switch_coup,'Name','WALL_E|KENN_E|BERI_E|ROSH_E',regex=True)
all_switch_coup_11kV_EXL5=UsefulPandas.filter_df_single(all_switch_coup_11kV_EXL5,'Grid','WALL_E|KENN_E|BERI_E|ROSH_E',regex=True)
all_switch_coup_11kV_EXL5=all_switch_coup_11kV_EXL5.sort_values(by='Name',ascending=True)
print(all_switch_coup_11kV_EXL5)

all_switch_coup_11kV_EXL5.to_csv(os.path.join(PSAfolder,'all_switch_coup_11kV_EXL5.csv'),index=False, header=True)

all_switches=pf_object.get_all_circuit_breakers_switch()
show3 = all_switches.head(5)
print(show3)

all_switches_COLO_33kV=UsefulPandas.filter_df_single(all_switches,'Grid','Cowley Local',regex=True)
print(all_switches_COLO_33kV)

print(all_switches_COLO_33kV[all_switches_COLO_33kV['Network_element_type'].isna()])

all_switches_COLO_33kV_lines_trafos=UsefulPandas.filter_df_single(all_switches_COLO_33kV,'Network_element_type','line|transformer',regex=True)
print(all_switches_COLO_33kV_lines_trafos)

all_switches_COLO_33kV_lines_trafos.to_csv(os.path.join(PSAfolder,'all_switches_COLO_33kV_lines_trafos.csv'),index=False, header=True)