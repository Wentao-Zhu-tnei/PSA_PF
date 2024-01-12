import sys
import os
import pandas as pd
import numpy as np
import UsefulPrinting
import PSA_SND_Utilities as ut
import PSA_SND_Constants as const
import PF_Config as pfconf
import UsefulPandas


dict_config = ut.PSA_SND_read_config(const.strConfigFile)
print(f'[const.strPFUser]=[{dict_config[const.strPFUser]}]')
print(f'[const.strPFPath]=[{dict_config[const.strPFPath]}]')


print(f'Using this python environment: {sys.executable}')
print(f'Using this python version: {sys.version}')

if os.path.exists(dict_config[const.strPFPath]):
    sys.path.append(dict_config[const.strPFPath])
elif os.path.exists(r"c:\Support\Python\3.9"): 
    sys.path.append(r"c:\Support\Python\3.9")
else:
    print('#####################################')
    print('Module powerfactory_interface.py : Power Factory Libraries not Available')
    
import powerfactory as pf

PSAstatus = const.PSAok
PSAerrMsg = "Power Factory __init"

print('Using PowerFactory version: ', pf.__version__)


class PowerFactoryClass(object):
    """
    This class launches PowerFactory and has several methods

    """

    def __init__(self, user_name=dict_config[const.strPFUser], multiProcess=False, ShowApp=False):
        global PSAstatus
        global PSAerrMsg
        print(f'Initialising PowerFactory with [ShowApp]={ShowApp} and [user_name]={user_name}')
        PSAstatus = const.PSAok
        PSAerrMsg = "Power Factory __init"
        if multiProcess:
            print("Multi-process PowerFactory")
            fnum = 0
            found = False
            while not found:
                workspace = dict_config[const.strWorkingFolder] + "\\tempPF\\workspace" + str(fnum)
                print("PF workspace = " + workspace)
                if not os.path.isdir(workspace):
                    os.makedirs(workspace)
                    found = True
                else:
                    fnum = fnum + 1
        
            self.app = pf.GetApplicationExt(None, None, f'/general.workspaceDirectory "{workspace}"')
        else:
            print("Single process PowerFactory")
            #self.app = 999
            self.app = pf.GetApplication(user_name)
        
        if str(self.app) == "None":
            print("Power Factory Error: No licences available")
            PSAstatus = const.PSAPFError
            PSAerrMsg = "Power Factory Error: No licences available"
        else:
            print(f'PowerFactory initialised correctly with user: [{user_name}]')

        if PSAstatus == const.PSAok:
            ## Activate a local cache so not every change will result in a database transaction
            #print("Activate local cache 2")
            #self.enableLocalCache()
            ## Deactivate the stop button in the toolbar
            self.app.SetUserBreakEnabled(0)
            ## Deactivate the graphics update
            self.app.SetGuiUpdateEnabled(0)
            if ShowApp == True:
                self.app.Show()
                self.app.ClearOutputWindow()
                self.app.PrintInfo('Hi there, PowerFactory started from Python script...')
                print('Running PowerFactory as an app on the side (SLOW)')
            else:
                print('Running PowerFactory in background (FAST)')

            self.user = self.app.GetCurrentUser()
        return
        ####################################################

    def get_all_projects(self):
        projects = self.user.GetContents('*.IntPrj', 0)
        return [project.loc_name for project in projects]
        ####################################################

    def check_project_exists(self, project_name, all_projects):

        project_name_noext = project_name.split('.')[0]

        if project_name_noext in all_projects:
            print(f'Project: [{project_name_noext}] already exists')
            project_exists = True


        else:
            print(f'Project: [{project_name_noext}] does not exists')
            project_exists = False

        return project_name_noext, project_exists
        ####################################################

    def import_project(self, project, main_dir):

        UsefulPrinting.print_double_spacer()

        all_projects = self.get_all_projects()

        # project_name_noext, project_exists = self.check_project_exists(project, all_projects)

        project_exists = False
        if project_exists:
            # print(f'Project: [{project}] already exists, need to activate only')
            print(f'Project: [{project}] exists, need to activate only...')
        else:
            #print(f"Deleting any existing projects...")
            # oProject = self.user.GetContents('*.IntPrj', 0)[0]
            lProjects = self.user.GetContents('.IntPrj', 0)
            for oProject in lProjects:
                if oProject.loc_name == project.split(".")[0]:
                    oProject.Deactivate()
                    oProject.Delete()
                    #print(f"Existing project has been deleted...")

            print(f'Importing project [{project}] into PowerFactory')
            import_proj = self.user.CreateObject('CompfdImport', 'Import')
            ### project_name should include the extension: XXXX.pfd
            import_proj.SetAttribute("e:g_file",
                                     os.path.join(main_dir, project))
            location = self.user

            import_proj.g_target = location

            import_proj.activatePrj = 1

            ierr = import_proj.Execute()

            if ierr == 0:
                ## Activate a local cache so not every change will result in a database transaction
                #print("Activate local cache 3")
                #self.enableLocalCache()
                ## Deactivate the stop button in the toolbar
                self.app.SetUserBreakEnabled(0)
                ## Deactivate the graphics update
                self.app.SetGuiUpdateEnabled(0)
                #self.app.PrintInfo("Import command returns no error")
            else:
                self.app.PrintError("ERROR IN Import command")
                print("ERROR IN Import command")
                return ierr

        UsefulPrinting.print_double_spacer()

        return None
        ####################################################

    def activate_project(self, project_name):
        ### project_name should include the extension: XXXX.pfd
        project = self.app.ActivateProject(project_name)
        #print(f'Activated project: [{project_name}]')
        # self.app.PrintPlain(f'Activated project: {project_name}')
        intPrj = self.app.GetActiveProject()
        intPrj.Purge()
        #print("Project purged")
        self.app.ClearRecycleBin()
        #print("Recycle bin emptied")
        ## Activate a local cache so not every change will result in a database transaction
        #print("Activate local cache 1")
        #self.enableLocalCache()
        UsefulPrinting.print_double_spacer()
        return
        ####################################################

    def purge_project(self):
        intPrj = self.app.GetActiveProject()
        intPrj.Purge()
        #print("Exit: Project purged")
        self.app.ClearRecycleBin()
        #print("Exit: Recycle bin emptied")
        return
        ####################################################

    def get_all_study_cases(self):

        self.study_case_folder = self.app.GetProjectFolder('study')
        # self.app.PrintPlain(self.study_case_folder)
        study_cases_raw = self.study_case_folder.GetContents('*.IntCase')
        self.study_cases = [case.GetFullName().split('\\')[-1].split('.')[0] for case in study_cases_raw]

        for case in study_cases_raw:
            print(case.GetFullName().split('\\')[-1].split('.')[0])
            # self.app.PrintPlain(case)

        return self.study_case_folder
        ####################################################

    def activate_study_case(self, study_case_name):

        self.active_study_case = self.study_case_folder.GetContents(study_case_name + '.IntCase')[0]
        self.active_study_case.Activate()
        self.app.PrintInfo('Activating operation scenario: ' + study_case_name)
        print(f'Activating study case: {study_case_name}')
        # self.app.PrintPlain(f"Activating study case:{study_case_name}")

        return
        ####################################################

    def get_all_operation_scenarios(self):

        operation_scenario_folder = self.app.GetProjectFolder('scen')
        # self.app.PrintPlain(operation_scenario_folder)
        operation_scenarios_raw = operation_scenario_folder.GetContents()
        self.operation_scenarios = [scenario.GetFullName().split('\\')[-1] for scenario in operation_scenarios_raw]

        for scenario in operation_scenarios_raw:
            print(scenario.GetFullName().split('\\')[-1].split('.')[0])
            # self.app.PrintPlain(scenario)

        return operation_scenario_folder
        ####################################################

    def get_active_operation_scenario(self):

        operation_scenario_obj = self.app.GetActiveScenario()
        operation_scenario_name = operation_scenario_obj.GetFullName().split('\\')[-1].split('.')[0]
        print(f'\t \t - Active Operation Scenario: {operation_scenario_name}')
        # self.app.PrintPlain(f"Active Operation Scenario: {operation_scenario_obj}")

        return operation_scenario_obj
        ####################################################

    def activate_operation_scenario(self, operational_scenario):

        operation_scenario_folder = self.app.GetProjectFolder('scen')
        active_operation_scenario = operation_scenario_folder.GetContents(operational_scenario + '.IntScenario')[0]

        active_operation_scenario.Activate()
        self.app.PrintInfo('Activating operation scenario: ' + operational_scenario)
        print(f'Activating operation scenario: {operational_scenario}')

        return active_operation_scenario
        ####################################################

    def get_transformers_2w_results(self, *args):

        if len(args) != 0:
            specific_trafos = [arg for arg in args]
            print(specific_trafos)

            # Collect all relevant transformers
            transformers = self.app.GetCalcRelevantObjects(specific_trafos[0] + '*.ElmTr2')

        else:

            # Collect all relevant transformers
            transformers = self.app.GetCalcRelevantObjects('*.ElmTr2')

        if pfconf.bDebug:
            print('Number of transformers 2-winding: ', len(transformers))

        transformers_df = pd.DataFrame(columns=['PF_object', 'Name', 'Folder', 'Grid', 'Substation',
                                                'To_HV', 'From_LV',
                                                'From_HV_brk', 'To_LV_brk',
                                                'From_HV_status', 'To_LV_status',
                                                'Min_tap', 'Max_tap',
                                                'Auto_tap', 'V_setpoint', 'V_lb', 'V_ub',
                                                'Type', 'HV', 'LV', 'Vector_group', 'I_rated_kA_LV', 'I_nom_act_kA_LV',
                                                'S_rated', 'S_rated_act', 'No_load_loss', 'Copper_loss',
                                                'SC_voltage_pos', 'SC_voltage_neg', 'V_HV', 'V_LV',
                                                'Tap_position_actual',
                                                'I_HV', 'I_LV', 'I_max_act', 'P_HV', 'Q_HV', 'S_HV', 'P_LV', 'Q_LV',
                                                'S_LV', 'S_max_act',
                                                'Ploss', 'Qloss', 'Loading', 'Overload', 'OutService', 'Lat', 'Lon',
                                                'CIM_ID'])

        for transformer, index in zip(transformers, range(len(transformers))):

            name = transformer.loc_name

            type_id = transformer.typ_id
            type_id = None if type_id is None else transformer.typ_id.loc_name

            grid = transformer.cpGrid.loc_name
            folder = transformer.fold_id.loc_name

            tx_bushv = transformer.bushv
            if tx_bushv is None:
                tx_bushv = None
            else:
                tx_bushv = transformer.bushv.cterm.loc_name

            tx_buslv = transformer.buslv
            if tx_buslv == None:
                tx_buslv = None
            else:
                tx_buslv = transformer.buslv.cterm.loc_name

            ## Get the cubicles at each end of the transformer
            try:
                transformer_bus1_cub = transformer.GetCubicle(0)
            except AttributeError:
                transformer_bus1_cub = None

            try:
                transformer_bus2_cub = transformer.GetCubicle(1)
            except AttributeError:
                transformer_bus2_cub = None

            if transformer_bus1_cub != None:
                ## Get the switches
                try:
                    transformer_bus1_cub_switch = transformer_bus1_cub.GetChildren(0)[0]
                    transformer_bus1_cub_switch_name = transformer_bus1_cub_switch.loc_name
                    ## Get the status of the switches
                    transformer_bus1_cub_switch_isclosed = bool(transformer_bus1_cub_switch.isclosed)

                    if transformer_bus1_cub_switch_isclosed == True:
                        transformer_bus1_cub_switch_isclosed = 'CLOSED'
                    else:
                        transformer_bus1_cub_switch_isclosed = 'OPEN'

                ## If there are no switches then fill with None
                except IndexError:
                    transformer_bus1_cub_switch = None
                    transformer_bus1_cub_switch_isclosed = None

            if transformer_bus2_cub != None:
                try:
                    transformer_bus2_cub_switch = transformer_bus2_cub.GetChildren(0)[0]
                    transformer_bus2_cub_switch_name = transformer_bus2_cub_switch.loc_name
                    transformer_bus2_cub_switch_isclosed = bool(transformer_bus2_cub_switch.isclosed)

                    if transformer_bus2_cub_switch_isclosed == True:
                        transformer_bus2_cub_switch_isclosed = 'CLOSED'
                    else:
                        transformer_bus2_cub_switch_isclosed = 'OPEN'

                ## If there are no switches then fill with None
                except IndexError:
                    transformer_bus2_cub_switch = None
                    transformer_bus2_cub_switch_isclosed = None

            from_terminal_status = transformer_bus1_cub_switch_isclosed
            to_terminal_status = transformer_bus2_cub_switch_isclosed

            if type_id is None:

                HV = None
                LV = None
                vecgrp = None
                Copper_loss = None
                No_load_loss = None
                SC_voltage_pos = None
                SC_voltage_zero = None

            else:

                HV = transformer.typ_id.utrn_h
                LV = transformer.typ_id.utrn_l
                vecgrp = transformer.typ_id.vecgrp
                Copper_loss = transformer.typ_id.pcutr
                No_load_loss = transformer.typ_id.pfe
                SC_voltage_pos = transformer.typ_id.uktr
                SC_voltage_zero = transformer.typ_id.uk0tr

            if HV > 100:
                substation_type = 'BSP'
            elif HV < 100 and HV > 30:
                substation_type = 'Primary'
            elif HV < 30 and HV > 10:
                substation_type = 'Secondary'

            #    transformerHV=transformer.bushvn
            S_rated = transformer.Snom
            S_rated_act = transformer.Snom_a

            I_rated_LV = transformer.GetAttribute('e:Inom_l')
            I_rated_act_LV = transformer.GetAttribute('e:Inom_l_a')

            ## Tap positions
            min_tx_tap = transformer.optapmin
            max_tx_tap = transformer.optapmax
            auto_tap = str(bool(transformer.ntrcn))

            if auto_tap == 'True':
                V_setpoint = transformer.usetp
                V_low = transformer.usp_low
                V_up = transformer.usp_up

            else:
                V_setpoint = None
                V_low = None
                V_up = None

            # Positive/zero sequence impedance

            transformer_out = str(bool(transformer.outserv))

            transformer_lat = transformer.GPSlat
            transformer_lon = transformer.GPSlon

            transformer_rdfid = transformer.cimRdfId

            if len(transformer_rdfid) > 0:
                transformer_rdfid = transformer_rdfid[0]

            if transformer_out == 'False':

                ### Check if the transformer has results
                iHasResults = transformer.HasResults()

                if iHasResults == 1:

                    ## HV
                    V_HV = transformer.GetAttribute('m:u:bushv')
                    P_HV = transformer.GetAttribute('m:P:bushv')
                    Q_HV = transformer.GetAttribute('m:Q:bushv')
                    S_HV = transformer.GetAttribute('m:S:bushv')
                    I_HV = transformer.GetAttribute('m:I:bushv')

                    ## LV
                    V_LV = transformer.GetAttribute('m:u:buslv')
                    P_LV = transformer.GetAttribute('m:P:buslv')
                    Q_LV = transformer.GetAttribute('m:Q:buslv')
                    S_LV = transformer.GetAttribute('m:S:buslv')
                    I_LV = transformer.GetAttribute('m:I:buslv')

                    ## Get max apparent power and max current
                    I_max = max(I_LV, I_HV)
                    S_max = max(S_HV, S_LV)
                    loading = transformer.GetAttribute('c:loading')
                    Ploss = transformer.GetAttribute('c:Ploss')
                    Qloss = transformer.GetAttribute('c:Qloss')

                    ## Tap position
                    tap_position = transformer.GetAttribute('c:nntap')
                    if loading > 95:
                        overloaded = True
                    else:
                        overloaded = False

                else:

                    V_HV = None
                    V_LV = None
                    tap_position = None
                    I_LV = None
                    I_HV = None
                    I_max = None
                    P_HV = None
                    Q_HV = None
                    S_HV = None
                    P_LV = None
                    Q_LV = None
                    S_LV = None
                    S_max = None
                    loading = None
                    Ploss = None
                    Qloss = None
                    overloaded = None

                transformers_df.loc[index] = [transformer, name, folder, grid, substation_type,
                                              tx_bushv, tx_buslv,
                                              transformer_bus1_cub_switch_name, transformer_bus2_cub_switch_name,
                                              from_terminal_status, to_terminal_status,
                                              min_tx_tap, max_tx_tap,
                                              auto_tap, V_setpoint, V_low, V_up,
                                              type_id, HV, LV, vecgrp, I_rated_LV, I_rated_act_LV,
                                              S_rated, S_rated_act, No_load_loss, Copper_loss,
                                              SC_voltage_pos, SC_voltage_zero, V_HV, V_LV, tap_position,
                                              I_HV, I_LV, I_max, P_HV, Q_HV, S_HV, P_LV, Q_LV, S_LV, S_max,
                                              Ploss, Qloss, loading, overloaded, transformer_out,
                                              transformer_lat, transformer_lon, transformer_rdfid]

            else:
                if pfconf.bDebug:
                    print(f'Transformer [{name}] is out of service')
                V_HV = 0
                V_LV = 0
                tap_position = 0
                I_LV = 0
                I_HV = 0
                I_max = 0
                P_HV = 0
                Q_HV = 0
                S_HV = 0
                P_LV = 0
                Q_LV = 0
                S_LV = 0
                S_max = 0
                loading = 0
                Ploss = 0
                Qloss = 0
                overloaded = 0
                transformers_df.loc[index] = [transformer, name, folder, grid, substation_type,
                                              tx_bushv, tx_buslv,
                                              transformer_bus1_cub_switch_name, transformer_bus2_cub_switch_name,
                                              from_terminal_status, to_terminal_status,
                                              min_tx_tap, max_tx_tap,
                                              auto_tap, V_setpoint, V_low, V_up,
                                              type_id, HV, LV, vecgrp, I_rated_LV, I_rated_act_LV,
                                              S_rated, S_rated_act, No_load_loss, Copper_loss,
                                              SC_voltage_pos, SC_voltage_zero, V_HV, V_LV, tap_position,
                                              I_HV, I_LV, I_max, P_HV, Q_HV, S_HV, P_LV, Q_LV, S_LV, S_max,
                                              Ploss, Qloss, loading, overloaded, transformer_out,
                                              transformer_lat, transformer_lon, transformer_rdfid]

        return transformers_df

    ####################################################

    # def get_transformers_2w_results(self, *args):
    #
    #     if len(args) != 0:
    #         specific_trafos = [arg for arg in args]
    #         print(specific_trafos)
    #
    #         # Collect all relevant transformers
    #         transformers = self.app.GetCalcRelevantObjects(specific_trafos[0] + '*.ElmTr2')
    #
    #     else:
    #
    #         # Collect all relevant transformers
    #         transformers = self.app.GetCalcRelevantObjects('*.ElmTr2')
    #
    #     print('Total number of transformers 2-winding in network model: ', len(transformers))
    #
    #     transformers_df = pd.DataFrame(columns=['Name', 'Folder', 'Grid',
    #                                             'To_HV', 'From_LV', 'Min_tap', 'Max_tap',
    #                                             'Auto_tap', 'V_setpoint', 'V_lb', 'V_ub',
    #                                             'Type', 'HV', 'LV', 'Vector_group',
    #                                             'S_rated', 'S_rated_act', 'No_load_loss', 'Copper_loss',
    #                                             'SC_voltage_pos', 'SC_voltage_neg', 'V_HV', 'V_LV',
    #                                             'Tap_position_actual',
    #                                             'Current_LV', 'P_HV', 'Q_HV', 'S_HV', 'P_LV', 'Q_LV', 'S_LV',
    #                                             'S_max_act',
    #                                             'Ploss', 'Qloss', 'Loading', 'Overload', 'OutService', 'Lat', 'Lon',
    #                                             'CIM_ID'])
    #
    #     for transformer, index in zip(transformers, range(len(transformers))):
    #
    #         name = transformer.loc_name
    #
    #         type_id = transformer.typ_id
    #         if type_id == None:
    #             type_id = None
    #         else:
    #             type_id = transformer.typ_id.loc_name
    #
    #         grid = transformer.cpGrid.loc_name
    #         folder = transformer.fold_id.loc_name
    #
    #         tx_bushv = transformer.bushv
    #         if tx_bushv == None:
    #             tx_bushv = None
    #         else:
    #             tx_bushv = transformer.bushv.cterm.loc_name
    #
    #         tx_buslv = transformer.buslv
    #         if tx_buslv == None:
    #             tx_buslv = None
    #         else:
    #             tx_buslv = transformer.buslv.cterm.loc_name
    #
    #         if type_id == None:
    #
    #             HV = None
    #             LV = None
    #             vecgrp = None
    #             Copper_loss = None
    #             No_load_loss = None
    #             SC_voltage_pos = None
    #             SC_voltage_zero = None
    #
    #         else:
    #
    #             HV = transformer.typ_id.utrn_h
    #             LV = transformer.typ_id.utrn_l
    #             vecgrp = transformer.typ_id.vecgrp
    #             Copper_loss = transformer.typ_id.pcutr
    #             No_load_loss = transformer.typ_id.pfe
    #             SC_voltage_pos = transformer.typ_id.uktr
    #             SC_voltage_zero = transformer.typ_id.uk0tr
    #
    #         #    transformerHV=transformer.bushvn
    #         S_rated = transformer.Snom
    #         S_rated_act = transformer.Snom_a
    #
    #         ## Tap positions
    #         min_tx_tap = transformer.optapmin
    #         max_tx_tap = transformer.optapmax
    #         auto_tap = str(bool(transformer.ntrcn))
    #
    #         if auto_tap == 'True':
    #             V_setpoint = transformer.usetp
    #             V_low = transformer.usp_low
    #             V_up = transformer.usp_up
    #
    #         else:
    #             V_setpoint = None
    #             V_low = None
    #             V_up = None
    #
    #         # Positive/zero sequence impedance
    #
    #         transformer_out = str(bool(transformer.outserv))
    #
    #         transformer_lat = transformer.GPSlat
    #         transformer_lon = transformer.GPSlon
    #
    #         transformer_rdfid = transformer.cimRdfId
    #
    #         if len(transformer_rdfid) > 0:
    #             transformer_rdfid = transformer_rdfid[0]
    #
    #         if transformer_out == 'False':
    #
    #             ### Check if the line has results
    #             iHasResults = transformer.HasResults()
    #
    #             if iHasResults == 1:
    #
    #                 V_HV = transformer.GetAttribute('m:u:bushv')
    #                 V_LV = transformer.GetAttribute('m:u:buslv')
    #                 tap_position = transformer.GetAttribute('c:nntap')
    #                 Current_LV = transformer.GetAttribute('m:I:1')
    #                 P_HV = transformer.GetAttribute('m:P:0')
    #                 Q_HV = transformer.GetAttribute('m:Q:0')
    #                 S_HV = transformer.GetAttribute('m:S:bushv')
    #                 P_LV = transformer.GetAttribute('m:P:1')
    #                 Q_LV = transformer.GetAttribute('m:Q:1')
    #                 S_LV = transformer.GetAttribute('m:S:buslv')
    #                 S_max = max(S_HV, S_LV)
    #                 loading = transformer.GetAttribute('c:loading')
    #                 #                Ploss=P_HV+P_LV
    #                 Ploss = transformer.GetAttribute('c:Ploss')
    #                 Qloss = transformer.GetAttribute('c:Qloss')
    #
    #                 if loading > 95:
    #                     overloaded = True
    #                 else:
    #                     overloaded = False
    #
    #             else:
    #
    #                 V_HV = None
    #                 V_LV = None
    #                 tap_position = None
    #                 Current_LV = None
    #                 P_HV = None
    #                 Q_HV = None
    #                 S_HV = None
    #                 P_LV = None
    #                 Q_LV = None
    #                 S_LV = None
    #                 S_max = None
    #                 loading = None
    #                 Ploss = None
    #                 Qloss = None
    #                 overloaded = None
    #
    #             transformers_df.loc[index] = [name, folder, grid, tx_bushv, tx_buslv, min_tx_tap, max_tx_tap,
    #                                           auto_tap, V_setpoint, V_low, V_up,
    #                                           type_id, HV, LV, vecgrp,
    #                                           S_rated, S_rated_act, No_load_loss, Copper_loss,
    #                                           SC_voltage_pos, SC_voltage_zero, V_HV, V_LV, tap_position,
    #                                           Current_LV, P_HV, Q_HV, S_HV, P_LV, Q_LV, S_LV, S_max,
    #                                           Ploss, Qloss, loading, overloaded, transformer_out,
    #                                           transformer_lat, transformer_lon, transformer_rdfid]
    #
    #         else:
    #             print(f'Transformer [{name}] is out of service')
    #
    #     return transformers_df

    def get_lines_results(self, *args):

        if len(args) != 0:
            specific_lines = [arg for arg in args]
            print(specific_lines)
            # Collect all relevant lines
            lines = self.app.GetCalcRelevantObjects(specific_lines[0] + '*.ElmLne')

        else:
            # Collect all relevant lines
            lines = self.app.GetCalcRelevantObjects('*.ElmLne')

        print('Number of lines: ', len(lines))

        # Start empty dataframe to store results as we loop over lines
        lines_df = pd.DataFrame(columns=['Name', 'Type', 'Grid', 'Voltage_level',
                                         'Cable_OHL', 'Description', 'Line_phases',
                                         'From', 'To',
                                         'From_brk', 'To_brk',
                                         'From_status', 'To_status',
                                         'Length', 'V_nom', 'I_rated_kA', 'I_nom_act_kA',
                                         'R0', 'X0', 'R1', 'X1', 'B0', 'B1',
                                         'R0_km', 'X0_km', 'R1_km', 'X1_km',
                                         'P_HV', 'Q_HV', 'P_LV', 'Q_LV',
                                         'V_LV', 'I_LV', 'I_HV', 'I_max_act',
                                         'V_HV', 'Loading', 'Ploss', 'Qloss', 'Overload', 'OutService', 'CIM_ID'])

        cable_ohl_dict = {0: 'cable', 1: 'oh_line'}

        for line, index in zip(lines, range(len(lines))):

            # Parameters
            name = line.loc_name
            # print(name)
            grid = line.cpGrid.loc_name
            type_line = line.typ_id.loc_name

            cable_ohl = cable_ohl_dict[line.typ_id.cohl_]
            description_type = line.desc

            if len(description_type) > 0:
                description_type = description_type[0]

            line_rdfid = line.cimRdfId

            if len(line_rdfid) > 0:
                line_rdfid = line_rdfid[0]

            from_terminal = line.bus1
            if from_terminal == None:
                from_terminal = str(None)
            else:
                from_terminal = line.bus1.cterm.loc_name

            to_terminal = line.bus2
            if to_terminal == None:
                to_terminal = str(None)
            else:
                to_terminal = line.bus2.cterm.loc_name

            ## Get the cubicles
            line_bus1_cub = line.GetCubicle(0)
            line_bus2_cub = line.GetCubicle(1)

            ## Get the switches
            # line_bus1_cub_switch=line_bus1_cub.GetChildren(0)[0]
            # line_bus2_cub_switch=line_bus2_cub.GetChildren(0)[0]

            ### If there are not switches within the cubicle then return None
            try:
                line_bus1_cub_switch = line_bus1_cub.GetChildren(0)[0]
            except IndexError:
                line_bus1_cub_switch = None

            try:
                line_bus2_cub_switch = line_bus2_cub.GetChildren(0)[0]
            except IndexError:
                line_bus2_cub_switch = None

            if line_bus1_cub_switch != None:
                ## Get the name of the switch
                line_bus1_cub_switch_name = line_bus1_cub_switch.loc_name
                ## Get the status of the switches
                line_bus1_cub_switch_isclosed = bool(line_bus1_cub_switch.isclosed)

                ## Get the status of the switch and convert to CLOSED/OPEN string
                if line_bus1_cub_switch_isclosed == True:
                    line_bus1_cub_switch_isclosed = 'CLOSED'
                else:
                    line_bus1_cub_switch_isclosed = 'OPEN'

            else:
                line_bus1_cub_switch_name = None
                line_bus1_cub_switch_isclosed = None

            if line_bus2_cub_switch != None:
                ## Get the name of the switch
                line_bus2_cub_switch_name = line_bus2_cub_switch.loc_name
                ## Get the status of the switch
                line_bus2_cub_switch_isclosed = bool(line_bus2_cub_switch.isclosed)

                ## Get the status of the switch and convert to CLOSED/OPEN string
                if line_bus2_cub_switch_isclosed == True:
                    line_bus2_cub_switch_isclosed = 'CLOSED'
                else:
                    line_bus2_cub_switch_isclosed = 'OPEN'

            else:
                line_bus2_cub_switch_name = None
                line_bus2_cub_switch_isclosed = None

            from_terminal_status = line_bus1_cub_switch_isclosed
            to_terminal_status = line_bus2_cub_switch_isclosed

            length = line.dline
            rated_voltage = line.Unom
            rated_current = line.typ_id.sline
            line_phases = line.typ_id.nlnph
            I_nom_act = line.Inom_a
            R0, X0 = line.R0, line.X0
            R1, X1 = line.R1, line.X1
            B0, B1 = line.B0, line.B1

            R1_length = line.typ_id.rline  ## pos-seq R per length (ohm/km)
            X1_length = line.typ_id.xline  ## pos-seq R per length (ohm/km)
            R0_length = line.typ_id.rline0  ## pos-seq R per length (ohm/km)
            X0_length = line.typ_id.rline0  ## pos-seq R per length (ohm/km)

            if rated_voltage > 100:
                line_volt_type = '132kV'
            elif rated_voltage < 100 and rated_voltage > 30:
                line_volt_type = '33kV'
            elif rated_voltage < 30 and rated_voltage > 10:
                line_volt_type = '11kV'
            elif rated_voltage < 10 and rated_voltage > 0:
                line_volt_type = '0.4kV'

            # Out of service?
            line_out = str(bool(line.outserv))

            if line_out == 'False':

                ### Check if the line has results

                iHasResults = line.HasResults()

                if iHasResults == 1:

                    ## Measurements if there are results available

                    ## HV
                    P_HV = line.GetAttribute('m:P:bus1')
                    Q_HV = line.GetAttribute('m:Q:bus1')
                    S_HV = line.GetAttribute('m:S:bus1')
                    I_HV = line.GetAttribute('m:I:bus1')
                    V_HV = line.GetAttribute('m:u:bus1')

                    ## LV
                    P_LV = line.GetAttribute('m:P:bus2')
                    Q_LV = line.GetAttribute('m:Q:bus2')
                    S_LV = line.GetAttribute('m:S:bus2')
                    I_LV = line.GetAttribute('m:I:bus2')
                    V_LV = line.GetAttribute('m:u:bus2')

                    Current_max = max(I_LV, I_HV)

                    loading = line.GetAttribute('c:loading')
                    P_Loss = line.GetAttribute('c:Ploss')
                    Q_Loss = line.GetAttribute('c:Qloss')

                    if loading > 95:
                        overloaded = True
                    else:
                        overloaded = False

                else:

                    ## If no results are available then use None

                    P_HV = None
                    Q_HV = None
                    P_LV = None
                    Q_LV = None
                    V_LV = None
                    I_LV = None
                    I_HV = None
                    Current_max = None
                    V_HV = None
                    loading = None
                    P_Loss = None
                    Q_Loss = None
                    overloaded = None

                lines_df.loc[index] = [name, type_line, grid, line_volt_type,
                                       cable_ohl, description_type, line_phases,
                                       from_terminal, to_terminal,
                                       line_bus1_cub_switch_name, line_bus2_cub_switch_name,
                                       from_terminal_status, to_terminal_status,
                                       length, rated_voltage, rated_current, I_nom_act, R0, X0, R1, X1, B0, B1,
                                       R1_length, X1_length, R0_length, X0_length,
                                       P_HV, Q_HV, P_LV, Q_LV, V_LV, I_LV, I_HV, Current_max,
                                       V_HV, loading, P_Loss, Q_Loss, overloaded, line_out, line_rdfid]

            else:
                print(f'Line [{name}] is out of service')

        return lines_df

    ####################################################

    def get_loads_results(self):

        # Collect all relevant loads
        loads = self.app.GetCalcRelevantObjects('*.ElmLod')
        if pfconf.bDebug:
            print('Total number of loads in network model: ', len(loads))

        loads_df = pd.DataFrame(columns=['Name', 'Grid', 'Cubicle', 'Terminal',
                                         'Distance_buses', 'Distance_km',
                                         'Terminal_voltage', 'Connection',
                                         'Unbalanced', 'OutService',
                                         'Load_transformer', 'P_nom', 'Q_nom', 'V_nom',
                                         'P_actual', 'Q_actual', 'V_actual',
                                         'Overvoltage', 'Undervoltage', 'CIM_ID',
                                         'PF_object'])  ### modified to include the PF object as part of the df

        for load, index in zip(loads, range(len(loads))):

            load_name = load.loc_name
            P, Q, V = load.plini, load.qlini, load.u0
            load_grid = str(load.cpGrid.loc_name)
            load_cubicle = load.bus1.loc_name
            load_terminal = load.bus1.cterm
            load_terminal_name = load_terminal.loc_name
            load_terminal_voltage = load.bus1.cterm.uknom
            load_conn = load.phtech
            load_unbalanced = bool(int(load.i_sym))  # 0: Balanced, 1: Unbalanced

            load_out = str(bool(load.outserv))

            load_trafo = load.iLoadTrf

            load_rdf_id = load.cimRdfId

            if len(load_rdf_id) > 0:
                load_rdf_id = load_rdf_id[0]
            else:
                load_rdf_id = 'None'

            # Measurements
            ### Check if the line has results

            iHasResults = load.HasResults()

            if iHasResults == 1:

                P_actual_LV = load.GetAttribute('m:P:bus1')
                Q_actual_LV = load.GetAttribute('m:Q:bus1')
                voltage_actual_LV = load.GetAttribute('m:u1:bus1')

                if voltage_actual_LV > 1.05:
                    Over_voltage = 1
                    Under_voltage = None
                elif voltage_actual_LV < 0.95:
                    Under_voltage = 1
                    Over_voltage = None
                else:
                    Over_voltage = 0
                    Under_voltage = 0

                terminal_dist_bus = load_terminal.GetAttribute('e:ciDist')
                terminal_dist_km = load_terminal.GetAttribute('b:dist')

            else:
                # print('Need to run a calculation before obtaining results!')
                P_actual_LV = None
                Q_actual_LV = None
                voltage_actual_LV = None
                Under_voltage = None
                Over_voltage = None

                terminal_dist_bus = None
                terminal_dist_km = None

            loads_df.loc[index] = [load_name, load_grid, load_cubicle, load_terminal_name,
                                   terminal_dist_bus, terminal_dist_km,
                                   load_terminal_voltage, load_conn,
                                   load_unbalanced, load_out,
                                   load_trafo, P, Q, V, P_actual_LV, Q_actual_LV, voltage_actual_LV,
                                   Over_voltage, Under_voltage, load_rdf_id, load]

        return loads_df

    ####################################################

    def get_loads_names_grid_obj(self):

        # Collect all relevant loads
        loads = self.app.GetCalcRelevantObjects('*.ElmLod')
        if pfconf.bDebug:
            print('Total number of loads in network model: ', len(loads))

        ### modified to include the PF object as part of the df
        loads_df = pd.DataFrame(columns=['Name', 'Grid', 'Name_grid', 'PF_object'])

        for load, index in zip(loads, range(len(loads))):
            load_name = load.loc_name
            load_grid = str(load.cpGrid.loc_name)
            load_name_grid = f'{load_name}_{load_grid}'

            loads_df.loc[index] = [load_name, load_grid, load_name_grid, load]
        loads_df["PF_object_str"] = loads_df["PF_object"].apply(str)

        return loads_df

    ####################################################

    def set_load_pq_setpoints(self, load, Pset, Qset, verbose=False):

        # Collect all relevant loads
        specific_load = load + '.ElmLod'
        specific_load = self.app.GetCalcRelevantObjects(specific_load)[0]

        specific_load.plini = Pset
        specific_load.qlini = Qset

        if verbose == True:
            print(f'Changing P/Q setpoint for load: {load}')

        return specific_load

    ####################################################

    def prepare_loadflow(self, ldf_mode, algorithm, trafo_tap, shunt_tap, zip_load, q_limits, phase_shift,
                         trafo_tap_limit, max_iter):

        ldf_mode_dict = {'balanced': 0, 'unbalanced': 1, 'dc': 2}
        algorithm_dict = {'current': 0, 'power': 1}
        trafo_tap_dict = {'no': 0, 'yes': 1}
        shunt_tap_dict = {'no': 0, 'yes': 1}
        zip_load_dict = {'no': 0, 'yes': 1}
        q_limits_dict = {'no': 0, 'yes': 1}
        phase_shift_dict = {'no': 0, 'yes': 1}
        trafo_tap_limit_dict = {'no': 0, 'yes': 1}

        self.ldf = self.app.GetFromStudyCase('ComLdf')
        # self.app.PrintPlain(self.ldf)

        self.ldf.SetAttribute('i_power', algorithm_dict[algorithm])

        self.ldf.SetAttribute('iopt_at', trafo_tap_dict[trafo_tap])
        self.ldf.SetAttribute('iopt_asht', shunt_tap_dict[shunt_tap])
        self.ldf.SetAttribute('iopt_net', ldf_mode_dict[ldf_mode])
        self.ldf.SetAttribute('iopt_pq', zip_load_dict[zip_load])
        self.ldf.SetAttribute('iopt_lim', q_limits_dict[q_limits])
        self.ldf.SetAttribute('iPST_at', phase_shift_dict[phase_shift])
        self.ldf.SetAttribute('iopt_optaplim', trafo_tap_limit_dict[trafo_tap_limit])
        self.ldf.SetAttribute('itrlx', max_iter)
        self.ldf.SetAttribute('iopt_check', 0)  ### Output report
        self.ldf.SetAttribute('iKeepCalc', 1)  ### Output results even calculation fails

        # print(
        #     f'Output report with parameters:\n loadmax: [{self.ldf.loadmax:.3f}],\n vlmin: [{self.ldf.vlmin:.3f}], vlmax: [{self.ldf.vlmax:.3f}]')
        # print(
        #     f'Preparing **[{ldf_mode}]** Loadflow with:\n auto_trafo_tap: [{trafo_tap}],\n and auto_shunt_tap: [{shunt_tap}],\n and load dependency: [{zip_load}],\n and q limits: [{q_limits}],\n and phase shifters tap: [{phase_shift}]')
        # print(f'Considering the operational limits for tap changers: [{trafo_tap_limit}]')
        # print(f'Maximum number of iterations: [{max_iter}]')
        UsefulPrinting.print_double_spacer()

    ####################################################

    def enableLocalCache(self):
        #stop writing to DB and save in local memory
        self.app.SetWriteCacheEnabled(1)

    def disableLocalCache(self):
        #restart writing to DB and update DB
        self.app.SetWriteCacheEnabled(0)

    def updatePFdb(self):
        #force changes to be written to DB
        self.app.WriteChangesToDb()

    def run_loadflow(self):
        # print('Running Loadflow')
        ierr = self.ldf.Execute()

        if ierr == 0:
            self.app.PrintInfo("Load Flow command returns no error")
            # print("Load Flow command returns no error")
        else:
            self.app.PrintError("ERROR IN LOAD FLOW")
            print("ERROR IN LOAD FLOW")

        return ierr

    ####################################################

    def get_terminal_conn_elements(self, specific_terminal):

        terminal = self.app.GetCalcRelevantObjects(specific_terminal + '.ElmTerm')[0]

        terminal_name = terminal.loc_name

        terminal_elements = terminal.GetConnectedElements()

        terminal_elements_df = pd.DataFrame(columns=['Name', 'Class',
                                                     'Full_name', 'Terminal',
                                                     'Terminal_from', 'Terminal_to',
                                                     'Terminal_from_voltage', 'Terminal_to_voltage',
                                                     'CIM_ID', 'Folder',
                                                     'NPhase', 'Phase_info', 'PF_object'])

        ### Loop through the elements connected to this terminal and get parameters
        for idx, element in enumerate(terminal_elements):

            element_name = element.loc_name
            element_class = element.GetClassName()
            if len(element.cimRdfId) > 0:
                element_rdf_id = element.cimRdfId[0]
            else:
                element_rdf_id = None
            element_folder = element.fold_id.loc_name

            element_full_name = str(f'{element_name}.{element_class}')
            element_full_name_folder = str(element)

            if element_class == 'ElmTr2':
                element_nphase = element.bushv.nphase
                element_phase_info = element.bushv.cPhInfo
                element_terminal_from = element.bushv.cterm.loc_name
                element_terminal_from_voltage = element.bushv.cterm.uknom
                element_terminal_to = element.buslv.cterm.loc_name
                element_terminal_to_voltage = element.buslv.cterm.uknom


            elif element_class == 'ElmLne':
                element_nphase = element.bus1.nphase
                element_phase_info = element.bus1.cPhInfo
                element_terminal_from = element.bus1.cterm.loc_name
                element_terminal_from_voltage = element.bus1.cterm.uknom
                element_terminal_to = element.bus2.cterm.loc_name
                element_terminal_to_voltage = element.bus2.cterm.uknom

            else:
                element_nphase = element.bus1.nphase
                element_phase_info = element.bus1.cPhInfo
                element_terminal_from = element.bus1.cterm.loc_name
                element_terminal_from_voltage = element.bus1.cterm.uknom
                element_terminal_to = None
                element_terminal_to_voltage = None

            # print(f'Element {idx}: {element_name}, class: {element_class}')

            terminal_elements_df.loc[idx] = [element_name, element_class,
                                             element_full_name, terminal_name,
                                             element_terminal_from, element_terminal_to,
                                             element_terminal_from_voltage, element_terminal_to_voltage,
                                             element_rdf_id, element_folder,
                                             element_nphase, element_phase_info, element]

        return terminal_elements_df

    def get_line_terminal_result(self, specific_line, lne_threshold):

        # Collect all relevant lines
        line = self.app.GetCalcRelevantObjects(specific_line + '.ElmLne')[0]

        # Parameters
        name = line.loc_name
        grid = line.cpGrid.loc_name
        type_line = line.typ_id.loc_name

        line_df = pd.DataFrame(columns=['PF_Object', 'Name',
                                        'Type', 'Grid',
                                        'From', 'To',
                                        "From_Status", "To_Status",
                                        'P_HV', 'Q_HV', 'S_HV', 'I_HV', 'V_HV',
                                        'P_LV', 'Q_LV', 'S_LV', 'I_LV', 'V_LV',
                                        'I_nom_act_kA_lv', 'I_nom_act_kA_hv',
                                        'I_max_act',
                                        'loading_threshold_pct',
                                        'Loading', 'loading_diff_perc',
                                        'Ploss', 'Qloss', 'Overload',
                                        'OutService', 'CIM_ID', 'Service'])

        line_rdfid = line.cimRdfId

        line_Inom = line.Inom_a

        if len(line_rdfid) > 0:
            line_rdfid = line_rdfid[0]

        from_terminal = line.bus1
        if from_terminal is None:
            from_terminal = str(None)
        else:
            from_terminal = line.bus1.cterm.loc_name

        to_terminal = line.bus2
        if to_terminal is None:
            to_terminal = str(None)
        else:
            to_terminal = line.bus2.cterm.loc_name

        ## Get the cubicles
        line_bus1_cub = line.GetCubicle(0)
        line_bus2_cub = line.GetCubicle(1)

        ## Get the switches
        line_bus1_cub_switch = line_bus1_cub.GetChildren(0)[0]
        line_bus2_cub_switch = line_bus2_cub.GetChildren(0)[0]

        ## Get the status of the switches
        line_bus1_cub_switch_isclosed = bool(line_bus1_cub_switch.isclosed)
        line_bus2_cub_switch_isclosed = bool(line_bus2_cub_switch.isclosed)

        from_terminal_status = line_bus1_cub_switch_isclosed
        to_terminal_status = line_bus2_cub_switch_isclosed

        length = line.dline
        rated_voltage = line.Unom

        # Out of service?
        line_out = str(bool(line.outserv))

        if line_out == 'False':

            ### Check if the line has results

            iHasResults = line.HasResults()

            if iHasResults == 1:

                # Measurements

                ## HV
                P_HV = line.GetAttribute('m:P:bus1')
                Q_HV = line.GetAttribute('m:Q:bus1')
                S_HV = line.GetAttribute('m:S:bus1')
                I_HV = line.GetAttribute('m:I:bus1')
                V_HV = line.GetAttribute('m:u:bus1') * rated_voltage

                ## LV
                P_LV = line.GetAttribute('m:P:bus2')
                Q_LV = line.GetAttribute('m:Q:bus2')
                S_LV = line.GetAttribute('m:S:bus2')
                I_LV = line.GetAttribute('m:I:bus2')
                V_LV = line.GetAttribute('m:u:bus2') * rated_voltage

                Current_LV = line.GetAttribute('m:I:bus1')
                Current_HV = line.GetAttribute('m:I:bus2')
                Current_max = max(Current_LV, Current_HV)

                loading = line.GetAttribute('c:loading')
                P_Loss = line.GetAttribute('c:Ploss')
                Q_Loss = line.GetAttribute('c:Qloss')

                if loading > lne_threshold:
                    overloaded = True
                else:
                    overloaded = False

                loading_diff_perc = UsefulPandas.get_perc_change(loading, lne_threshold)

                if I_HV > I_LV:
                    line_service = "import"
                else:
                    line_service = "export"
            else:

                P_HV = None
                Q_HV = None
                S_HV = None
                I_HV = None
                V_HV = None
                P_LV = None
                Q_LV = None
                S_LV = None
                I_LV = None
                V_LV = None
                Current_LV = None
                Current_HV = None
                Current_max = None
                loading = None
                P_Loss = None
                Q_Loss = None
                overloaded = None
                line_service = None
            line_df.loc[0] = [line, name,
                              type_line, grid,
                              from_terminal, to_terminal,
                              from_terminal_status, to_terminal_status,
                              P_HV, Q_HV, S_HV, I_HV, V_HV,
                              P_LV, Q_LV, S_LV, I_LV, V_LV,
                              line_Inom, line_Inom,
                              Current_max,
                              lne_threshold,
                              loading, loading_diff_perc,
                              P_Loss, Q_Loss,
                              overloaded, line_out, line_rdfid, line_service]

        else:

            print(f'Line [{name}] is out of service')

        return line_df
        ####################################################

    def get_transformer_terminal_result(self, specific_transformer, transformer_threshold):

        # Collect all relevant lines
        transformer = self.app.GetCalcRelevantObjects(specific_transformer + '*.ElmTr2')[0]

        transformers_df = pd.DataFrame(columns=['PF_object', 'Name',
                                                'Type', 'Grid',
                                                'To_HV', 'From_LV',
                                                'From_HV_Status', 'To_HV_Status',
                                                'P_HV', 'Q_HV', 'S_HV', 'I_HV', 'V_HV',
                                                'P_LV', 'Q_LV', 'S_LV', 'I_LV', 'V_LV',
                                                'I_nom_act_kA_lv', 'I_nom_act_kA_hv', 
                                                'I_max_act',
                                                'loading_threshold_pct',
                                                'Loading', 'loading_diff_perc',
                                                'Ploss', 'Qloss', 'Overload',
                                                'OutService', 'CIM_ID','Service'])

        name = transformer.loc_name

        type_id = transformer.typ_id
        type_id = None if type_id is None else transformer.typ_id.loc_name

        grid = transformer.cpGrid.loc_name

        tx_bushv = transformer.bushv
        if tx_bushv is None:
            tx_bushv = None
        else:
            tx_bushv = transformer.bushv.cterm.loc_name

        tx_buslv = transformer.buslv
        if tx_buslv == None:
            tx_buslv = None
        else:
            tx_buslv = transformer.buslv.cterm.loc_name

        if type_id is None:
            HV = None
            LV = None
        else:
            HV = transformer.typ_id.utrn_h
            LV = transformer.typ_id.utrn_l

        S_rated = transformer.Snom
        S_rated_act = transformer.Snom_a

        I_rated_LV = transformer.GetAttribute('e:Inom_l')
        I_rated_act_LV = transformer.GetAttribute('e:Inom_l_a')
        I_rated_act_HV = transformer.GetAttribute('e:Inom_h_a')

        ## Get the cubicles
        transformer_buslv_cub = transformer.GetCubicle(0)
        transformer_bushv_cub = transformer.GetCubicle(1)

        ## Get the switches
        transformer_buslv_cub_switch = transformer_buslv_cub.GetChildren(0)[0]
        transformer_bushv_cub_switch = transformer_bushv_cub.GetChildren(0)[0]

        ## Get the status of the switches
        transformer_buslv_cub_switch_isclosed = bool(transformer_buslv_cub_switch.isclosed)
        transformer_bushv_cub_switch_isclosed = bool(transformer_bushv_cub_switch.isclosed)

        tx_buslv_status = transformer_buslv_cub_switch_isclosed
        tx_bushv_status = transformer_bushv_cub_switch_isclosed

        transformer_out = str(bool(transformer.outserv))

        transformer_rdfid = transformer.cimRdfId

        if len(transformer_rdfid) > 0:
            transformer_rdfid = transformer_rdfid[0]

        if transformer_out == 'False':

            ### Check if the line has results
            iHasResults = transformer.HasResults()

            if iHasResults == 1:

                ## HV
                V_HV = transformer.GetAttribute('m:u:bushv') * HV
                P_HV = transformer.GetAttribute('m:P:bushv')
                Q_HV = transformer.GetAttribute('m:Q:bushv')
                S_HV = transformer.GetAttribute('m:S:bushv')
                I_HV = transformer.GetAttribute('m:I:bushv')

                ## LV
                V_LV = transformer.GetAttribute('m:u:buslv') * LV
                P_LV = transformer.GetAttribute('m:P:buslv')
                Q_LV = transformer.GetAttribute('m:Q:buslv')
                S_LV = transformer.GetAttribute('m:S:buslv')
                I_LV = transformer.GetAttribute('m:I:buslv')

                ## Get max apparent power and max current
                I_max = max(I_LV, I_HV)
                S_max = max(S_HV, S_LV)
                loading = transformer.GetAttribute('c:loading')
                Ploss = transformer.GetAttribute('c:Ploss')
                Qloss = transformer.GetAttribute('c:Qloss')

                ## Tap position
                tap_position = transformer.GetAttribute('c:nntap')

                if loading > transformer_threshold:
                    overloaded = True
                else:
                    overloaded = False

                loading_diff_perc = UsefulPandas.get_perc_change(loading, transformer_threshold)

                if P_HV>0:
                    transformer_service = "import"
                else:
                    transformer_service = "export"
            else:

                V_HV = None
                V_LV = None
                I_LV = None
                I_HV = None
                I_max = None
                P_HV = None
                Q_HV = None
                S_HV = None
                P_LV = None
                Q_LV = None
                S_LV = None
                S_max = None
                loading = None
                Ploss = None
                Qloss = None
                overloaded = None
                loading_diff_perc = None
                transformer_service = None
            transformers_df.loc[0] = [transformer, name, type_id, grid,
                                      tx_bushv, tx_buslv,
                                      tx_bushv_status, tx_buslv_status,
                                      P_HV, Q_HV, S_HV, I_HV, V_HV,
                                      P_LV, Q_LV, S_LV, I_LV, V_LV,
                                      I_rated_act_LV, I_rated_act_HV,
                                      I_max,
                                      transformer_threshold,
                                      loading, loading_diff_perc,
                                      Ploss, Qloss,
                                      overloaded,
                                      transformer_out, transformer_rdfid, transformer_service]

        else:
            print(f'Transformer [{name}] is out of service')
            # V_HV = 0
            # V_LV = 0
            # I_LV = 0
            # I_HV = 0
            # I_max = 0
            # P_HV = 0
            # Q_HV = 0
            # S_HV = 0
            # P_LV = 0
            # Q_LV = 0
            # S_LV = 0
            # S_max = 0
            # loading = 0
            # Ploss = 0
            # Qloss = 0
            # overloaded = 0
            # loading_diff_perc = 0
            #
            # transformers_df.loc[0] = [transformer, name, type_id, grid,
            #                           tx_bushv, tx_buslv,
            #                           tx_bushv_status, tx_buslv_status,
            #                           P_HV, Q_HV, S_HV, I_HV, V_HV,
            #                           P_LV, Q_LV, S_LV, I_LV, V_LV,
            #                           I_rated_act_LV, I_max,
            #                           transformer_threshold,
            #                           loading, loading_diff_perc,
            #                           Ploss, Qloss,
            #                           overloaded,
            #                           transformer_out, transformer_rdfid]

        return transformers_df

    def get_line_results(self, lAttr, lHeader):
        """
        To access required line parameters
        """

        lines = self.app.GetCalcRelevantObjects("*.ElmLne")
        line_results = []
        for line in lines:
            if line.GetAttribute("e:outserv") == 0:
                if line.HasResults() == 1:
                    line_results.append([line.GetAttribute(attr) for attr in lAttr])

        df = pd.DataFrame(data=np.array(line_results), columns=lHeader)

        return df

    def get_trafo_results(self, lAttr, lHeader):
        """
        To access required trafo parameters
        """
        trafos = self.app.GetCalcRelevantObjects("*.ElmTr")
        trafo_results = []
        for trafo in trafos:
            if trafo.GetAttribute("e:outserv") == 0:
                if trafo.HasResults() == 1:
                    trafo_results.append([trafo.GetAttribute(attr) for attr in lAttr])

        df = pd.DataFrame(data=np.array(trafo_results), columns=lHeader)

        return df

    def get_asset_results(self, type, bAllAssets, lSubAssetsNames, lAttr, lHeader):
        """
        To access required asset parameters
        """
        if type == const.strAssetLine:
            assets = self.app.GetCalcRelevantObjects("*.ElmLne")
        elif type == const.strAssetTrans:
            assets = self.app.GetCalcRelevantObjects("*.ElmTr2")

        lAllAssetsNames = []
        for asset in assets:
            if asset.HasResults() == 1:
                lAllAssetsNames.append(asset.loc_name)

        if bAllAssets:
            lSubAssetsNames = lAllAssetsNames

        lAssetResults = []
        asset_results = []
        for asset in assets:
            if asset.HasResults() == 1:
                if asset.loc_name in lSubAssetsNames:
                    for attr in lAttr:
                        if attr == "e:cpGrid":
                            strAssetResults = str(
                                os.path.basename(os.path.dirname(str(asset.GetAttribute(attr)))).split('.')[0])
                        elif attr == "e:cimRdfId":
                            if asset.GetAttribute(attr):
                                strAssetResults = str(asset.GetAttribute(attr)[0])
                            else:
                                strAssetResults = "Null"
                        elif attr == "e:outserv":
                            if asset.GetAttribute(attr) == 0:
                                strAssetResults = "IN"
                            else:
                                strAssetResults = "OUT"
                        elif attr == "m:P:bus1":
                            if type == const.strAssetLine:
                                strAssetResults = str(asset.GetAttribute(attr))
                            elif type == const.strAssetTrans:
                                strAssetResults = str(asset.GetAttribute("m:P:buslv"))
                        elif attr == "m:I:bus1":
                            if type == const.strAssetLine:
                                strAssetResults = str(asset.GetAttribute(attr))
                            elif type == const.strAssetTrans:
                                strAssetResults = str(asset.GetAttribute("m:I:buslv"))
                        elif attr == "m:I:bus2":
                            if type == const.strAssetLine:
                                strAssetResults = str(asset.GetAttribute(attr))
                            elif type == const.strAssetTrans:
                                strAssetResults = str(asset.GetAttribute("m:I:bushv"))
                        else:
                            strAssetResults = str(asset.GetAttribute(attr))
                        asset_results.append(strAssetResults)

                if bool(asset_results):
                    lAssetResults.append(asset_results)
                asset_results = []

        df = pd.DataFrame(data=lAssetResults, columns=lHeader)

        return df

    def get_sync_gen_pq_minmax(self, sync_gen, verbose=False):

        # Collect all relevant Generators
        specific_gens = sync_gen + '*.ElmSym'
        generator = self.app.GetCalcRelevantObjects(specific_gens)[0]

        generator_df = pd.DataFrame(columns=['Name', 'Type', 'Motor_mode', 'Grid',
                                             'OutService', 'P_max', 'P_min',
                                             'Q_max', 'Q_min', ])

        gen_name = generator.loc_name
        generator_out = str(bool(generator.outserv))
        gen_type = generator.GetClassName()
        generator_motor = bool(generator.i_mot)
        if generator_motor == False:
            generator_motor = 'generator'
        else:
            generator_motor = 'motor'
        gen_grid = generator.cpGrid.loc_name
        P, Q = generator.pgini, generator.qgini
        Pmax = generator.Pmax_uc
        Pmin = generator.Pmin_uc
        Pmax_rated = generator.P_max
        Pmax_ratingfactor = generator.pmaxratf
        Qmin, Qmax = generator.cQ_min, generator.cQ_max

        generator_df.loc[0] = [gen_name, gen_type, generator_motor, gen_grid,
                               generator_out, Pmax, Pmin, Qmax, Qmin]

        if verbose == True:
            print(
                f'Getting generator [{gen_name}], type: {gen_type}, Out: {generator_out}, P/Pmax: {P:.3f}/{Pmax:.3f}, Q/Qmax: {Q:.3f}/{Qmax:.3f}')

        return generator, P, Q, Pmax, Pmin, Qmin, Qmax, generator_df

    ####################################################

    def create_sync_gens_p_setpoints_steps(self, sync_gen, Pstart, steps, verbose=True):

        generator, P, Q, Pmax, Pmin, Qmin, Qmax, generator_df = self.get_sync_gen_pq_minmax(sync_gen, verbose)

        indv_step_P = Pmax * steps / 100

        endpoint = False

        if Pstart == 0:
            #### Using the numpy.linspace function which offers more options for floating point values
            # steps_P=list(np.arange(0, Pmax+indv_step_P, indv_step_P))
            steps_P = list(np.linspace(0, Pmax, steps, endpoint=endpoint))

        elif Pstart == 'Pmin':
            #### Using the numpy.linspace function which offers more options for floating point values
            # steps_P=list(np.arange(0, Pmax+indv_step_P, indv_step_P))
            steps_P = list(np.linspace(Pmin, Pmax, steps, endpoint=endpoint))

        if verbose == True:
            print(f'This is the list of P setpoints in {steps} steps')
            print(f'Pmax={Pmax}, indv_step_P:{indv_step_P}, endpoint incl?:{endpoint}')
            UsefulPrinting.print_list(steps_P)

        return generator, steps_P

    ####################################################

    def set_sync_gen_pq_minmax(self, sync_gen, mode_motor, Pmax, Pmin, verbose=False):

        # Collect all relevant Generators
        specific_gens = sync_gen + '*.ElmSym'
        generator = self.app.GetCalcRelevantObjects(specific_gens)[0]

        gen_name = generator.loc_name
        gen_type = generator.GetClassName()
        generator_out = str(bool(generator.outserv))

        if mode_motor == True:
            generator.i_mot = 1
        else:
            generator.i_mot = 0

        generator.Pmax_uc = Pmax
        generator.Pmin_uc = Pmin

        if verbose == True:
            print(
                f'Setting generator [{gen_name}], type: {gen_type}, Out: {generator_out}, Pmax: {Pmax:.3f}, Pmin: {Pmin:.3f}')

        return generator, Pmax, Pmin

    ####################################################

    def set_sync_gens_pq_setpoints(self, sync_gen, Pset, Qset, mode_motor=False, verbose=True):

        UsefulPrinting.print_spacer()

        generator, P, Q, Pmax, Pmin, Qmin, Qmax, generator_df = self.get_sync_gen_pq_minmax(sync_gen, verbose)
        generator_out = str(bool(generator.outserv))

        if mode_motor == True:
            generator.i_mot = 1
        else:
            generator.i_mot = 0

        generator.pgini = Pset
        generator.qgini = Qset

        if verbose == True:
            print(f'{generator.loc_name}, Out: {generator_out}, P: {generator.pgini:.3f}, Q: {generator.qgini:.3f}')

        UsefulPrinting.print_spacer()
        return generator

    ####################################################

    def flex_impact_wrt_network_element(self, flex_asset, network_element, Pstart, steps, lne_threshold,
                                        mode_motor=False, verbose=True):
        '''This function creates a dataframe with the impact of a flex asset on
            a network element when looping through a series of steps from 0 to Pmax
        '''

        ## Create a group of setpoints for a flex asset based on steps
        generator, steps_P = self.create_sync_gens_p_setpoints_steps(flex_asset, Pstart, steps, verbose)

        generator_name = generator.loc_name
        generator_motor = bool(generator.i_mot)
        if generator_motor == False:
            generator_motor = 'generator'
        else:
            generator_motor = 'motor'

        ## Create an empty dataframe to store results from run
        network_element_df_all_steps = pd.DataFrame(columns=['Name_network_element', 'Type', 'Grid',
                                                             'From', 'To', 'From_status', 'To_status',
                                                             'P_HV', 'Q_HV', 'S_HV', 'I_HV', 'V_HV',
                                                             'P_LV', 'Q_LV', 'S_LV', 'I_LV', 'V_LV',
                                                             'I_max_act', 'Loading', 'Ploss',
                                                             'Qloss', 'Overload', 'OutService', 'CIM_ID',
                                                             'Name_flex_asset', 'Generator_motor', 'Pflex'])

        # Collect all relevant Generators
        specific_network_element = network_element + '*.*'
        specific_network_element = self.app.GetCalcRelevantObjects(specific_network_element)[0]

        specific_network_element_name = specific_network_element.loc_name
        specific_network_element_type = specific_network_element.GetClassName()

        print(
            f'Getting the flex impact of {generator_name} wrt {specific_network_element_name}.{specific_network_element_type} with {steps} steps')
        print(f'Flex asset {generator_name} has Pmin: {steps_P[0]} and Pmax: {steps_P[-1]}')

        ## Iterate through the steps, assign to the flex asset and run load flow
        for idx, step in enumerate(steps_P):
            UsefulPrinting.print_spacer()
            print(f"[{idx}]-[{step}]")

            ### Change P/Q setpoints
            self.set_sync_gens_pq_setpoints(flex_asset, Pset=step, Qset=0, mode_motor=mode_motor, verbose=verbose)

            ### RUN LOAD FLOW
            self.run_loadflow()

            ## If the element is a Line then you need to get the results at the line terminals (bus1, bus2)
            if specific_network_element_type == 'ElmLne':
                network_element_df = self.get_line_terminal_result(network_element, lne_threshold).values

            ## If the element is a Transformer then you need to get the results at the transformer terminals (HV, LV)
            # elif specific_network_element_type == 'ElmTr2':
            #
            #     network_element_df = self.get_transformer_terminal_result(network_element, trafo_thrshold)

            ## Append the step to the np.array
            step = float(step)
            if generator_motor == 'generator':
                network_element_df = np.append(network_element_df, [generator_name, generator_motor])
                network_element_df = np.append(network_element_df, step)
            elif generator_motor == 'motor':
                network_element_df = np.append(network_element_df, [generator_name, generator_motor])
                network_element_df = np.append(network_element_df, step * (-1))

            network_element_df_all_steps.loc[idx] = network_element_df

        ## Reset the index of the dataframe after all the steps are completed
        network_element_df_all_steps = network_element_df_all_steps.reset_index(drop=True)

        return network_element_df_all_steps

    ####################################################

    def calculate_sensitivity_factors(self, flex_asset, network_element, flex_steps, Pstart, mode_motor, final_dir,
                                      verbose=True):

        UsefulPrinting.print_double_spacer()

        ## Create dataframe with the impact of a flex asset on a network element
        flex_net_all_steps_df = self.flex_impact_wrt_network_element(flex_asset, network_element, Pstart, flex_steps,
                                                                     mode_motor=mode_motor, verbose=verbose,
                                                                     lne_threshold=100)

        ## Get the direction of the flow on the network asset
        flex_net_all_steps_df[['P_flow_sign', 'P_flow_dir']] = flex_net_all_steps_df.apply(
            lambda x: UsefulPandas.get_direction_flow(x['P_LV'],
                                                      x['P_HV']), axis=1, result_type='expand')

        ## Get the sign of the power at the LV terminal of the network element
        flex_net_all_steps_df['P_LV_sign'] = flex_net_all_steps_df['P_LV'].apply(np.sign)

        ## Apply that sign to the MVA at the LV terminal of the network element
        flex_net_all_steps_df['S_LV_w_sign'] = flex_net_all_steps_df['S_LV'] * flex_net_all_steps_df['P_LV_sign']

        ## Get the rolling difference of the apparent power at the network element LV terminal (absolute)
        flex_net_all_steps_df['S_LV_diff_abs'] = UsefulPandas.running_diff_abs(
            np.array(flex_net_all_steps_df['S_LV_w_sign']), 0, 1)

        ## Get the rolling difference of the active power at the network element LV terminal (absolute)
        flex_net_all_steps_df['P_LV_diff_abs'] = UsefulPandas.running_diff_abs(np.array(flex_net_all_steps_df['P_LV']),
                                                                               0, 1)

        ## Get the rolling difference of the active power at the network element LV terminal (absolute)
        flex_net_all_steps_df['Pflex_diff_abs'] = UsefulPandas.running_diff_abs(
            np.array(flex_net_all_steps_df['Pflex']), 0, 1)

        ## Get the direction of the power flow difference in steps and the flow in P_LV_s0 at step 0
        flex_net_all_steps_df['PF_direction_opp'] = flex_net_all_steps_df.apply(
            lambda x: UsefulPandas.P_LV_diff_sign(x['P_LV_diff_abs'], flex_net_all_steps_df['P_LV'].loc[0]), axis=1)

        ## Get the direction of the power flow difference in steps and the flow in P_LV_sx at each step e.g. sx
        flex_net_all_steps_df['PF_direction_opp_rel_altern'] = flex_net_all_steps_df.apply(
            lambda x: UsefulPandas.P_LV_diff_sign(x['P_LV_diff_abs'], x['P_LV']), axis=1)

        ## Calculate the sensitivity factor of d_MVA_netasset/d_MW_flex

        ## Absolute step
        flex_net_all_steps_df['SF_abs'] = flex_net_all_steps_df.apply(
            lambda x: UsefulPandas.delta_MVA_delta_MW(x['S_LV_diff_abs'],
                                                      x['Pflex_diff_abs']),
            axis=1)

        ## Get the sensitivity factor direction, given by Delta_Pflow (flow in the network element) - absolute
        flex_net_all_steps_df['SF_abs_sign'] = flex_net_all_steps_df['P_LV_diff_abs'].apply(UsefulPandas.pd_npsign_alt)

        ## Validate the sensitivity factor by applying the sensitivity factor to the Pflex and add/subtract the initial S_LV with sign (S_LV_w_sign_s0)
        flex_net_all_steps_df['SF_validation'] = flex_net_all_steps_df.apply(lambda x: UsefulPandas.SF_validation(
            x['SF_abs'],
            x['SF_abs_sign'],
            x['PF_direction_opp'],
            flex_net_all_steps_df['P_LV'].loc[0],
            x['Pflex'],
            flex_net_all_steps_df['S_LV_w_sign'].loc[0]), axis=1)

        flex_net_all_steps_df['SF_validation_check'] = flex_net_all_steps_df['SF_validation'] == flex_net_all_steps_df[
            'S_LV']

        flex_net_all_steps_df = flex_net_all_steps_df.rename_axis("Step")

        ## Choose columns to be used for the final csv file
        selected_cols = ['Name_network_element', 'From', 'To',
                         'P_HV', 'S_HV', 'I_HV',
                         'P_LV', 'S_LV', 'I_LV',
                         'I_max_act', 'Loading',
                         'Overload', 'OutService',
                         'Name_flex_asset', 'Generator_motor',
                         'Pflex', 'P_flow_sign', 'P_flow_dir',
                         'P_LV_sign', 'S_LV_w_sign',
                         'S_LV_diff_abs', 'P_LV_diff_abs',
                         'Pflex_diff_abs', 'PF_direction_opp',
                         'SF_abs', 'SF_abs_sign',
                         'SF_validation', 'SF_validation_check']

        ## Save the dataframe as a csv file for offline analysis
        flex_net_all_steps_df[selected_cols].to_csv(os.path.join(final_dir,
                                                                 f'SF_sel_cols_{flex_asset}_wrt_{network_element}_steps_{str(flex_steps)}_Mot_mode_{mode_motor}.csv'),
                                                    index=True, header=True)

        UsefulPrinting.print_double_spacer()
        return flex_net_all_steps_df

    ####################################################

    def createGen(self, tgt_grid, name_gen, gen_terminal, gen):
        # self.study_case_object = self.app.GetProjectFolder('study')
        # self.scenario_object = self.app.GetProjectFolder('scen')
        networks = self.app.GetCalcRelevantObjects('*.ElmNet')
        for network in networks:
            if network.loc_name == tgt_grid:
                network.CreateObject('ElmSym', name_gen)
                self.new_gen = network.GetContents(name_gen + '.ElmSym')[0]
                self.new_gen.bus1 = self.app.GetCalcRelevantObjects(gen_terminal + '.ElmTerm')[0]
                self.new_gen.pgini = 10

            # self.studycase = self.app.GetActiveStudyCase()
        return

    def get_P_req_from_I_exceed_V_act_line(self, I_exceed, I_max_id, I_exceed_kA, V_act_kV):
        '''
        The voltage on the terminals of the line is the same so no need for I_max_id or for the extra elif condition
        '''
        if I_exceed == True:
            if I_max_id == "LV": # SEPM -> neagtive P_kw
                P_kW = -I_exceed_kA * V_act_kV * np.sqrt(3) * 1e3
            else: # SPM -> positive P_kw
                P_kW = I_exceed_kA * V_act_kV * np.sqrt(3) * 1e3
        else:
            P_kW = 0
        # if I_exceed == True:
        #     P_kW = I_exceed_kA * V_act_kV * np.sqrt(3) * 1e3
        # else:
        #     P_kW = I_exceed_kA * V_act_kV * np.sqrt(3) * 1e3
        # elif I_exceed == False:
        #     P_kW = I_exceed_kA * V_act_kV * np.sqrt(3) * 1e3
        # else:
        #     P_kW = 0
        return P_kW

    def get_lines_constraints(self, lines_df):
        '''Get the required active power injection to bring the current below the maximum rated current
        'I_nom_act_kA' is line.Inom_a
        'I_LV','I_HV' are line.GetAttribute('m:I:bus1') and line.GetAttribute('m:I:bus2')
        'V_LV', 'V_HV' are line.GetAttribute('m:u:bus1') and line.GetAttribute('m:u:bus2')
        'V_nom' is line.Unom
        '''
        lines_df['I_max_id'] = lines_df[['I_LV', 'I_HV', ]].abs().idxmax(axis=1)  # Always the terminal where constraint is calcualted
        lines_df['I_max_val'] = lines_df[['I_LV', 'I_HV', ]].abs().max(axis=1)  # kA

        ### this is where the I_nom_act_kA gets changed based on a maximum loading treshold (the purpose is to artificially create constraints)
        # lines_df['I_nom_act_kA_lv'] = lines_df['I_nom_act_kA_lv'] * (lines_df['loading_threshold_pct'] / 100)  # kA
        # lines_df['I_exceed'] = lines_df['I_max_val'] > lines_df['I_nom_act_kA_lv']  # kA
        # lines_df['I_exceed_val'] = lines_df['I_max_val'] - lines_df['I_nom_act_kA_lv']  # kA
        # I_max_val = I_lv  is the maximum for SEPM
        # I_max_val = I_HV is the maximum for SPM
        # I_max_id -> indicate whether SEPM/SPM
        if lines_df['I_max_id'] == 'I_LV':
            lines_df['I_nom_act_kA_lv'] = lines_df['I_nom_act_kA_lv'] * (lines_df['loading_threshold_pct'] / 100)  # kA
            lines_df['I_exceed'] = lines_df['I_max_val'] > lines_df['I_nom_act_kA_lv']  # kA
            lines_df['I_exceed_val'] = lines_df['I_max_val'] - lines_df['I_nom_act_kA_lv']  # kA
        elif lines_df['I_max_id'] == 'I_HV':
            lines_df['I_nom_act_kA_hv'] = lines_df['I_nom_act_kA_hv'] * (lines_df['loading_threshold_pct'] / 100)  # kA
            lines_df['I_exceed'] = lines_df['I_max_val'] > lines_df['I_nom_act_kA_hv']  # kA
            lines_df['I_exceed_val'] = lines_df['I_max_val'] - lines_df['I_nom_act_kA_hv']  # kA

        lines_df['P_req'] = lines_df.apply(lambda x: self.get_P_req_from_I_exceed_V_act_line(x['I_exceed'], x['I_max_id'],
                                                                                             x['I_exceed_val'],  # kA
                                                                                             x['V_LV']),  # kV
                                                                                             axis=1)

        return lines_df

    def get_P_req_from_I_exceed_V_act_trafo(self, P_max_id, I_exceed, I_exceed_kA, V_HV_act_kV, V_LV_act_kV, ):
        '''
        Should we use always I_LV regardless of which current is the maximum? 
        Given that we calculate required_power_kw at the point of the constraint?
        Should we output a P_kW regardless of an actual constraint? or P_kW=0?
        '''
        ##This is when there are constraints i.e. I_exceed == True
        if I_exceed == True:
            #### If I_max_id == 'I_LV' this is an export constraint because the flow is going from downstream to upstream (see trafo losses)
            if P_max_id == 'P_LV':
                print(f"Max_ID is LV: {P_max_id}")
                P_kW = -I_exceed_kA * V_LV_act_kV * np.sqrt(3) * 1e3 # SEPM -> negative P_kw
            #### else, this is an import constraint    
            else:
                print(f"Max_ID is HV: {P_max_id}")
                P_kW = I_exceed_kA * V_HV_act_kV * np.sqrt(3) * 1e3 # SPM -> positive P_kw

        ##This is when there are NO constraints i.e. I_exceed == False

        else:
            P_kW = 0

            # Alternative way of implementation for over procurement
            # if I_max_id == 'I_LV':
            #     P_kW = I_exceed_kA * V_LV_act_kV * np.sqrt(3) * 1e3
            # else:
            #     P_kW = I_exceed_kA * V_HV_act_kV * np.sqrt(3) * 1e3


        return P_kW

    def get_trafos_constraints(self, transformers_df):
        '''Get the required active power injection to bring the current below the maximum rated current
        'I_nom_act_kA_LV' is transformer.GetAttribute('e:Inom_l_a')
        'I_LV','I_HV' are transformer.GetAttribute('m:I:buslv') and transformer.GetAttribute('m:I:bushv')
        'V_LV', 'V_HV' are transformer.GetAttribute('m:u:buslv') and transformer.GetAttribute('m:u:bushv')
        'LV' and 'HV' are transformer.typ_id.utrn_l and transformer.typ_id.utrn_h
        '''
        # transformers_df['I_max_id'] = transformers_df[['I_LV', 'I_HV', ]].abs().idxmax(axis=1)  ## this needs to be the absolute value to get the residual/surplus constraint
        transformers_df['P_max_id'] = transformers_df[['P_HV', 'P_LV', ]].idxmax(axis=1)
        # transformers_df['I_max_val'] = transformers_df[['I_LV', 'I_HV', ]].abs().max(axis=1)
        
        if transformers_df['P_max_id'].values[0] == 'P_LV':
            ### this is where the I_nom_act_kA gets changed based on a maximum loading treshold (the purpose is to artificially create constraints)
            transformers_df['I_nom_act_kA_lv'] = transformers_df['I_nom_act_kA_lv'] * (
                        transformers_df['loading_threshold_pct'] / 100)

            transformers_df['I_exceed'] = transformers_df['I_LV'] > transformers_df['I_nom_act_kA_lv']

            transformers_df['I_exceed_val'] = transformers_df['I_LV'] - transformers_df['I_nom_act_kA_lv']  # should always use the LV side current?
        
        else:
            ### this is where the I_nom_act_kA gets changed based on a maximum loading treshold (the purpose is to artificially create constraints)
            transformers_df['I_nom_act_kA_hv'] = transformers_df['I_nom_act_kA_hv'] * (
                        transformers_df['loading_threshold_pct'] / 100)

            transformers_df['I_exceed'] = transformers_df['I_HV'] > transformers_df['I_nom_act_kA_hv']

            transformers_df['I_exceed_val'] = transformers_df['I_HV'] - transformers_df['I_nom_act_kA_hv']  # should always use the LV side current?


        transformers_df['P_req'] = transformers_df.apply(lambda x: self.get_P_req_from_I_exceed_V_act_trafo(x['P_max_id'],
                                                          x['I_exceed'],x['I_exceed_val'],x['V_HV'],x['V_LV']),axis=1)

        return transformers_df


    def toggle_switch_ind(self, switch_obj, action='toggle', verbose=False):
        '''Toggle on/off a specific switch'''

        # switch_obj_list = self.app.GetCalcRelevantObjects(f'{switch}.StaSwitch')
        # switch_obj = switch_obj_list[0]

        switch_obj_name = switch_obj.loc_name
        if verbose:
            print(f"the name of the switch is {switch_obj_name}")
        switch_obj_class = switch_obj.GetClassName()

        terminal = switch_obj.fold_id.cBusBar.loc_name

        network_element = switch_obj.fold_id.obj_id.loc_name

        ## Get status of the switch
        switch_obj_isclosed = bool(switch_obj.isclosed)
        if switch_obj_isclosed:
            switch_obj_isclosed_str = 'CLOSED'
        else:
            switch_obj_isclosed_str = 'OPEN'

        if action == 'open':
            switch_obj.on_off = 0
        elif action == 'close':
            switch_obj.on_off = 1
        elif action == 'toggle':
            ## Change switch status depending on status of isclosed
            switch_obj.on_off = int(not switch_obj_isclosed)
        else:
            print('No action taken')

        switch_obj_isclosed_new = bool(switch_obj.isclosed)

        if switch_obj_isclosed_new:
            switch_obj_isclosed_new_str = 'CLOSED'
        else:
            switch_obj_isclosed_new_str = 'OPEN'

        if verbose == True:
            UsefulPrinting.print_spacer()
            print(
                f'The status of the switch [{switch_obj_name}], inside terminal [{terminal}] of element [{network_element}], was changed: ')
            print('BEFORE')
            print(f'\tSwitch [{switch_obj_name}] was [{switch_obj_isclosed_str}]')
            print('AFTER')
            print(f'\tSwitch [{switch_obj_name}] now is [{switch_obj_isclosed_new_str}]')
            UsefulPrinting.print_spacer()

        return switch_obj

    ####################################################

    def get_all_circuit_breakers_coup(self):

        circuit_breakers = self.app.GetCalcRelevantObjects('*.ElmCoup')
        if pfconf.bDebug:
            print('Number of circuit breakers (ElmCoup): ', len(circuit_breakers))

        circuit_breakers_df = pd.DataFrame(
            columns=['Name', 'PF_Object_Coupler', 'Grid', 'Folder', 'From', 'To', 'Closed', 'On_Off', 'Phases', 'Type'])

        for circuit_breaker, index in zip(circuit_breakers, range(len(circuit_breakers))):
            dataobject = circuit_breaker

            name = circuit_breaker.loc_name

            grid = circuit_breaker.cpGrid.loc_name

            folder = circuit_breaker.fold_id.loc_name

            bus1 = circuit_breaker.bus1  # .cBusBar.loc_name
            if bus1 is None:
                bus1 = 'XX'
            else:
                bus1 = circuit_breaker.bus1.cBusBar.loc_name

            bus2 = circuit_breaker.bus2  # .cBusBar.loc_name
            if bus2 is None:
                bus2 = 'XX'
            else:
                bus2 = circuit_breaker.bus2.cBusBar.loc_name

            closed = bool(int(circuit_breaker.isclosed))
            on_off = bool(int(circuit_breaker.on_off))
            phases = int(circuit_breaker.nphase)
            cb_type = circuit_breaker.aUsage

            circuit_breakers_df.loc[index] = [name, dataobject, grid, folder,
                                              bus1, bus2,
                                              closed, on_off,
                                              phases, cb_type]

        return circuit_breakers_df

    ####################################################

    def get_all_circuit_breakers_switch(self):

        circuit_breakers = self.app.GetCalcRelevantObjects('*.StaSwitch')
        if pfconf.bDebug:
            print('Number of circuit breakers (StaSwitch): ', len(circuit_breakers))

        circuit_breakers_df = pd.DataFrame(columns=['Name', 'PF_Object_Switches', 'Grid', 'Folder',
                                                    'Terminal',
                                                    'Network_element',
                                                    'Network_element_type',
                                                    'Closed', 'On_Off',
                                                    'Type'])

        for index, circuit_breaker in enumerate(circuit_breakers):
            dataobject = circuit_breaker

            name = circuit_breaker.loc_name

            grid = circuit_breaker.cpGrid.loc_name

            folder = circuit_breaker.fold_id.loc_name

            terminal = circuit_breaker.fold_id.cBusBar.loc_name

            network_element = circuit_breaker.fold_id.obj_id.loc_name

            network_element_type = circuit_breaker.fold_id.obj_id.GetClassName()

            network_elements_dict = {
                'ElmLne': 'Line',
                'ElmTr2': 'Transformer',
                'ElmSym': 'Synchronous_generator',
                'ElmGenstat': 'Static_generator',
                'ElmXnet': 'Grid_external',
                'ElmShnt': 'Shunt_capacitor',
                'ElmLod': 'Load',
            }

            cb_type = circuit_breaker.aUsage

            closed = bool(int(circuit_breaker.isclosed))
            on_off = bool(int(circuit_breaker.on_off))

            circuit_breakers_df.loc[index] = [name, dataobject, grid, folder,
                                              terminal,
                                              network_element,
                                              network_element_type,
                                              closed, on_off,
                                              cb_type]

        ## Create a new column with the long name for the class based on PowerFactory naming convention
        circuit_breakers_df['Network_element_type'] = circuit_breakers_df['Network_element_type'].map(
            network_elements_dict)

        return circuit_breakers_df

    ####################################################

    def toggle_switch_coup_ind(self, switch_obj, action='toggle', verbose=False):
        '''Toggle on/off a specific switch from the ElmCoup type'''

        # switch_obj_list = self.app.GetCalcRelevantObjects(f'{switch}.ElmCoup')
        # switch_obj = switch_obj_list[0]

        switch_obj_name = switch_obj.loc_name
        if verbose:
            print(f"the name of the switch is {switch_obj_name}")
        switch_obj_class = switch_obj.GetClassName()

        from_terminal = switch_obj.bus1.cterm.loc_name
        to_terminal = switch_obj.bus2.cterm.loc_name

        ## Get status of the switch
        switch_obj_isclosed = bool(switch_obj.isclosed)
        if switch_obj_isclosed:
            switch_obj_isclosed_str = 'CLOSED'
        else:
            switch_obj_isclosed_str = 'OPEN'

        if action == 'open':
            switch_obj.on_off = 0
        elif action == 'close':
            switch_obj.on_off = 1
        elif action == 'toggle':
            ## Change switch status depending on status of isclosed
            switch_obj.on_off = int(not switch_obj_isclosed)
        else:
            print('No action taken')

        switch_obj_isclosed_new = bool(switch_obj.isclosed)

        if switch_obj_isclosed_new:
            switch_obj_isclosed_new_str = 'CLOSED'
        else:
            switch_obj_isclosed_new_str = 'OPEN'

        if verbose == True:
            UsefulPrinting.print_spacer()
            print(
                f'The status of the switch [{switch_obj_name}], going from [{from_terminal}] to [{to_terminal}], was changed: ')
            print('BEFORE')
            print(f'\tSwitch [{switch_obj_name}] was [{switch_obj_isclosed_str}]')
            print('AFTER')
            print(f'\tSwitch [{switch_obj_name}] now is [{switch_obj_isclosed_new_str}]')
            UsefulPrinting.print_spacer()

        return switch_obj
    ####################################################


def run(user_name, multiProcess):
    """to access powerfactory api calls"""
    opf = PowerFactoryClass(user_name=user_name, multiProcess=multiProcess, ShowApp=pfconf.bShowPFAPP)
    return opf, PSAstatus, PSAerrMsg

def activate_project(opf, project_name, main_dir):
    new_project = opf.import_project(project_name, main_dir)
    project_name = project_name[:-4]
    opf.activate_project(project_name)
    return


def PF_exit(opf):
    opf.purge_project()
    del opf
    print("Power Factory Exited")
    return