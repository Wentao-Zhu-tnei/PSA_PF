import papermill as pm
import PSA_SND_Constants as const
import PSA_analysis as pa
import os
from sys import argv
#import tkinter as tk
#import random
#from tkinter import ttk
#from tkinter.messagebox import showinfo, showerror
def run_digest(input_notebook, PSARunID, directory, scenario=None):
    if scenario == None:
        scenarios = [const.strBfolder,const.strMfolder,const.strMCfolder,const.strCfolder]
        scenarios = [scenario.split(' - ')[1][:-1] for scenario in scenarios if os.path.isdir(f'{directory}{scenario[:-1]}')]
    elif type(scenario) == list:
        scenarios = scenario
    else:
        scenarios = [scenario]

    save_fldr = f'{directory}\\'
    # Commented this next part out as nbconvert can't handle file paths with spaces in them, which \\7 - Analysis\\ has spaces in it
    #save_fldr = directory + const.strAnalysisFiles
    #if not os.path.isdir(save_fldr):
    #    os.mkdir(save_fldr)
    outputs = []
    for scenario in scenarios:
        print('########################')
        print(f'Executing for {scenario}')
        initials = pa.get_initials(scenario)
        
        output_notebook = f'{save_fldr}{initials}_Analysis_Digest_Output.ipynb'
        outputs.append(output_notebook)
        try:
            pm.execute_notebook(input_notebook,
                                output_notebook, 
                                parameters=dict(PSARunID=PSARunID, 
                                                directory=directory,  
                                                scenario=scenario)) 
                                #kernel_name='ttpsa')
            try:
                os.system('conda activate base')
                os.system(f'jupyter nbconvert --to html "{output_notebook}"')
                os.system('conda deactivate')
            except:
                os.system(f'jupyter nbconvert --to html "{output_notebook}"')
        except BaseException as err:
            print(err)

    print('Done\n')
    return outputs

#print(argv)
run_digest(argv[1], argv[2], argv[3], argv[4])