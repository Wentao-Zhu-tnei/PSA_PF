import powerfactory_interface as pf_interface
import PSA_SND_Utilities as ut
import PSA_SND_Constants as const
import os
import pandas as pd

target_day = 0
target_hour = 47
lAttr = ['b:loc_name', 'e:cpGrid', 'e:cimRdfId', "c:loading", "e:outserv"]
lHeader = ['Name', 'Grid', 'CIM_ID', 'Loading', 'IN/OUT_of_service']
pf_proj_dir = r"C:\Wentao Zhu\Project\PSA\Networks-Transition-Technical-Trials-PSA-1\data\input\COWLEY LOCAL BSP (20_07_22).pfd"

dict_config = ut.PSA_SND_read_config(const.strConfigFile)
pf_object = pf_interface.run(dict_config[const.strPFUser])
pf_interface.activate_project(pf_object, dict_config[const.strPFProject], pf_proj_dir)

PSA_folder = r"C:\Wentao Zhu\Project\PSA\Networks-Transition-Technical-Trials-PSA-1\data\results\MAN-2022-10-31-11-39"


# Load NeRDA data
NeRDA_fodler = PSA_folder + "/NeRDA_DATA"
for file in os.listdir(NeRDA_fodler):
    if "Line_Status" in file:
        line_name = file.split("-")[6]
        line_name = line_name.replace(".xlsx", "")
        file_path = NeRDA_fodler + "/" + file
        dfLineStatus = pd.read_excel(file_path,index_col=0, header=0)
        oLine = pf_object.app.GetCalcRelevantObjects(line_name + '.ElmLne')[0]
        oLine.outserv = int(dfLineStatus.iloc[target_hour][target_day])
    elif "Trafo_Status" in file:
        trafo_name = file.split("-")[6]
        trafo_name = trafo_name.replace(".xlsx", "")
        if trafo_name == "LONESOME FARM S-S":
            trafo_name = "LONESOME FARM S/S"
        if trafo_name == "KENNINGTON SEWAGE P-S":
            trafo_name = "KENNINGTON SEWAGE P/S"
        file_path = NeRDA_fodler + "/" + file
        dfTrafoStatus = pd.read_excel(file_path, index_col=0, header=0)
        oTrafo = pf_object.app.GetCalcRelevantObjects(trafo_name + '.ElmTr2')[0]
        oTrafo.outserv = int(dfTrafoStatus.iloc[target_hour][target_day])

# Load Generator data
SIA_fodler = PSA_folder + "/SIA_DATA"
for file in os.listdir(SIA_fodler):
    if file.startswith("SIA_GEN_"):
        df_gen = pd.read_excel(file_path)
        gen_name = file.split("-")[6]
        gen_name = gen_name.replace(".xlsx", "")
        oGen = pf_object.app.GetCalcRelevantObjects(gen_name+'.ElmSym')[0]
        oGen.pgini = df_gen.iloc[target_hour][target_day]

# Load load data
load_folder = PSA_folder + "/Downscaling Factor" + "/" + str(target_day) + "-" + str(target_hour)
for file in os.listdir(load_folder):
    if file.endswith(".xlsx"):
        df_load = pd.read_excel(load_folder + "/" + file, index_col=0, header=0)
        for load in list(df_load.index.values):
            oLoad = pf_object.app.GetCalcRelevantObjects(load + '.ElmLod')[0]
            oLoad.mode_inp = "SC"
            oLoad.slini = df_load.loc[load]["S"]
            oLoad.plini = df_load.loc[load]["P"]
            oLoad.qlini = df_load.loc[load]["Q"]

# Run power flow
pf_object.prepare_loadflow(ldf_mode=dict_config[const.strPFLFMode], algorithm=dict_config[const.strPFAlgrthm],
                     trafo_tap=dict_config[const.strPFTrafoTap], shunt_tap=dict_config[const.strPFShuntTap],
                     zip_load=dict_config[const.strPFZipLoad], q_limits=dict_config[const.strPFQLim],
                     phase_shift=dict_config[const.strPFPhaseShift],trafo_tap_limit=dict_config[const.strPFTrafoTapLim],
                     max_iter=int(dict_config[const.strPFMaxIter]))
ierr = pf_object.run_loadflow()

# Get Line & Trafo results
Debug_folder = PSA_folder + "\\" + "Debug"
if not os.path.isdir(Debug_folder):
    os.mkdir(Debug_folder)
lMonitoredLineAssets = ['Cable Segment_13369123']
df_line = pf_object.get_asset_results(type=const.strAssetLine, bAllAssets=False, lSubAssetsNames=lMonitoredLineAssets, lAttr=lAttr,
                                      lHeader=lHeader)
df_line.to_excel(Debug_folder + "Debug_Lines.xlsx")

lMonitoredTrafoAssets = ["BOWGRAVE COPSE"]
df_trafo = pf_object.get_asset_results(type=const.strAssetTrans, bAllAssets=False, lSubAssetsNames=lMonitoredTrafoAssets, lAttr=lAttr,
                                       lHeader=lHeader)
df_trafo.to_excel(Debug_folder + "Debug_Trafos.xlsx")

print("FINISHED!")







