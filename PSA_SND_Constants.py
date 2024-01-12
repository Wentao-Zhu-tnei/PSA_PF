### Definitions of constant variables for use in PSA modules ###

#User guide
strUserGuide = "PSA_User_Guide.pptx"

# Bool
bUTC = "PSA_SND_UTC"

# Int
PSAok = 0
PSAfileExistError = -1
PSAfileReadError = -2
PSAfileWriteError = -3
PSAfileTypeError = -5
PSAPFError = -4
PSAdataEntryError = -6
SF_default = -999

#Running modes
strRunModeMAN = "Manual"
strRunModeAUT = "Automatic"
strRunModeSEMIAUT = "Automatic_Simplified"

#Asset type names for input files
strAssetGen = "SyncMachine"
strAssetLine = "Line"
strAssetLoad = "Load"
strAssetTrans = "2wdgTransformer"
strAssetSwitch = "Switch"
strAssetCoupler = "Coupler"

#Asset type names for PF
strPFassetGen = "ElmSym"
strPFassetLine = "ElmLne"
strPFassetLoad = "ElmLod"
strPFassetTrans = "ElmTr2"
strPFassetSwitch = "StaSwitch"
strPFassetCoupler = "ElmCoup"


# Str
strINfolder = "\\0 - INPUT_DATA\\"
strSIAfolder = "\\1 - SIA_DATA\\"
strNewSIAfolder = "\\1 - SIA_DATA_NEW\\"
strNeRDAfolder = "\\2 - NeRDA_DATA\\"
strBfolder = "\\3 - BASE\\"
strMfolder = "\\4 - MAINT\\"
strCfolder = "\\5 - CONT\\"
strMCfolder = "\\6 - MAINT_CONT\\"
strAnalysisFiles = '\\7 - ANALYSIS\\'
strEventFiles = '\\8 - EVENT_DATA\\'
strSwitchFiles = '\\9 - SWITCH_DATA\\'
strDataInput = "\\data\\input\\"
strDataResults = "\\data\\results\\"
strExcelLogo = "\\packages\\psa\\Excel_Logo.png"
strTempSFfiles = "\\TEMP_SF_FILES\\"
strTempCRfiles = "\\TEMP_CR_FILES\\"

strFileDetectorFolder = "PSA_SND_FILE_DETECTOR_FOLDER"
strRunTimeParams = "-RUN TIME PARAMETERS.txt"
strPFUser = "PSA_SND_PF_USER"
strPFPath = "PSA_SND_PF_PATH"
# strPFProject = "PSA_SND_PF_MODEL"
strPFLFMode = "PSA_PF_LDF_MODE"
strPFAlgrthm = "PSA_PF_ALGRTHM"
strPFTrafoTap = "PSA_PF_TRAFO_TAP"
strPFShuntTap = "PSA_PF_SHUNT_TAP"
strPFZipLoad = "PSA_PF_ZIP_LOAD"
strPFQLim = "PSA_PF_Q_LIMITS"
strPFPhaseShift = "PSA_PF_PHASE_SHIFT"
strPFTrafoTapLim = "PSA_PF_TRAFO_TAP_LIMIT"
strPFMaxIter = "PSA_PF_MAX_ITER"
strPFStartDate = "PSA_PF_START_DATE"
strBSPModel = "BSP_MODEL"
strAssetData = "ASSET_DATA"
strMaint = "MAINTENANCE"
strMaintFile = "MAINT_FILE"
strCont = "CONTINGENCY"
strContFile = "CONT_FILE"
strbEvent = "EVENTS"
strEventFile = "EVENT_FILE"
strSwitch = "SWITCHING"
strSwitchFile = "SWITCH_FILE"
strPowFctr = "POWER_FACTOR"
strPowFctrLag = "LAGGING"
strAuto = "AUTOMATIC"
strRunMode = "RUNNING_MODE"
strDays = "DAYS"
strHalfHours = "HALF_HOURS"
strConfigFile = "PSA_SND_Config.txt"
strRefAssetFile = "AssetDataReference.xlsx"
strWorkingFolder = "PSA_SND_ROOT_WORKING_FOLDER"
# strInputFolder = "PSA_SND_INPUT_FOLDER"
# strResultsFolder = "PSA_SND_RESULTS_FOLDER"
strCalcFlexReqtsInterval = "PSA_SND_CALC_FLEX_REQTS_INTERVAL"
strHHMode = "PSA_SND_HH_MODE"
# strImageFile = "PSA_SND_EXCEL_LOGO"
strBSPDefault = "COWLEY LOCAL BSP 08_04_22"
strPrimaryDefault = "Rose Hill"
strNoFileDisp = "No file selected"
strNoFolderDisp = "No folder selected"
strDataSheet = "DATA"
strSIASheet = "SIA"
strAssetSheet = "ASSET"
strNeRDASheet = "NeRDA"
strSIAUsr = "PSA_SND_SIA_USER"
strSIAPwd = "PSA_SND_SIA_TOKEN"
strSIATimeOut = "PSA_SND_SIA_TIMEOUT"
strSIABaseURL = "PSA_SND_SIA_BASE_URL"
strSIALoginURL = "PSA_SND_SIA_LOGIN_URL"
strSIAURLFeeder = "PSA_SND_SIA_URL_FEEDER"
strSIAURLGroup = "PSA_SND_SIA_URL_GROUP"
strSIAURLGenerator = "PSA_SND_SIA_URL_GEN"
strSIAParams = "PSA_SND_SIA_PARAMS"
strNERDAUsr = "PSA_SND_NERDA_USER"
strNERDAPwd = "PSA_SND_NERDA_TOKEN"
strNERDAURL = "PSA_SND_NERDA_URL"
strNeRDATimeOut = "PSA_SND_NERDA_TIMEOUT"
strPSAFlexReqts = "_PSA_FLEX-REQTS"
strPSANoFlexReqts = "_PSA_NO-FLEX-REQTS"
strPSAConstraints = "_PSA_CONSTRAINTS"
strPSANoConstraints = "_PSA_NO-CONSTRAINTS"
strSNDResponses = "_SND_RESPONSES"
strPSAResponses = "_PSA_SF-RESPONSES"
strSNDContracts = "_SND_CONTRACTS"
strPSAContracts = "_PSA_SF-CONTRACTS"
strPSASFThreshold = "PSA_SF_THRESHOLD"
strSNDCandidateResponses = "_SND_CANDIDATE-RESPONSES"
strSNDCandidateContracts = "_SND_CANDIDATE-CONTRACTS"
strPSARejectedResponses = "_PSA_REJECTED-RESPONSES"
strPSARejectedContracts = "_PSA_REJECTED-CONTRACTS"
strPSAAcceptedResponses = "_PSA_ACCEPTED-RESPONSES"
strPSAAcceptedContracts = "_PSA_ACCEPTED-CONTRACTS"
# Lists
#lColAssetData = ["ASSET_ID","ASSET_TYPE","MAX_LOADING","RATING_1","RATING_2","RATING_3","EVENTS"]
lColAssetData = ["ASSET_ID","ASSET_TYPE","MAX_LOADING","BSP","PRIMARY","FEEDER"]
lColSIAData = ["ASSET_LONG_NAME","SIA_ID","ASSET_TYPE"]
lColNeRDAData = ["ASSET_LONG_NAME","NERDA_ID", "ASSET_TYPE", "SWITCHING"]
lColOutageData = ["ASSET_LONG_NAME", "ASSET_ID", "DURATION", "ASSET_TYPE", "START_OUTAGE", "END_OUTAGE", "DESCRIPTION"]
lColEventData = ["ASSET_LONG_NAME", "ASSET_ID", "DURATION", "ASSET_TYPE", "START_SERVICE", "END_SERVICE", "P_DISPATCH_KW", "DESCRIPTION"] #, "SCENARIO" #TB Proposed addition?
lColSwitchData = ["ASSET_LONG_NAME", "ASSET_ID", "ASSET_TYPE", "START_SERVICE", "END_SERVICE", "DURATION", "ACTION", "DESCRIPTION"] #, "SCENARIO" #TB Proposed addition?