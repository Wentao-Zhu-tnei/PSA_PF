import sys
import os
import PF_Config as pfconf
import PSA_PF_Functions as pffun
import PSA_SND_Utilities as ut
import PSA_SND_Constants as const
import powerfactory_interface

dict_config = ut.PSA_SND_read_config(const.strConfigFile)
os.environ["PATH"] = pfconf.PF_DIGSILENT_PATH + ";" + os.environ["PATH"]
sys.path.append(pfconf.PF_DIGSILENT_PATH)
import powerfactory_interface
app=powerfactory.GetApplication(dict_config[const.strPFUser])



script=app.GetCurrentScript()
importObj=script.CreateObject('CompfdImport','Import')

importObj.SetAttribute("e:g_file",r"C:\Users\...\projectToImport.pfd")

location=app.GetCurrentUser()
importObj.g_target=location

importObj.Execute()




dict_config = ut.PSA_SND_read_config(const.strConfigFile)
# sys.path.append(dict_config[const.strPFPath])
# main_dir=os.path.join(os.getcwd(),'PowerFactory')
# pf=PowerFactoryClass.PowerFactoryClass(user_name=dict_config[const.strPFUser],ShowApp=False)
# pf.import_project(dict_config[const.strPFProject],main_dir)
# project_name=dict_config[const.strPFProject][:-4]
#
#
# importObj.SetAttribute("e:g_file",r"C:\Wentao Zhu\Project\PSA")
#
#
#
# pf.activate_project(project_name)
#
#
#
#
#
# main_dir, pf = pffun.PF_activate_project(dict_config[const.strPFPath], dict_config[const.strPFUser],
#                                                      dict_config[const.strPFProject])




print(f'Using this python environment: {sys.executable}')
print(f'Using this python version: {sys.version}')

os.environ["PATH"] = pfconf.PF_DIGSILENT_PATH + ";" + os.environ["PATH"]
sys.path.append(pfconf.PF_DIGSILENT_PATH)

import powerfactory_interface as pfa

user =dict_config[const.strPFUser]
app=pfa.GetApplication(user)
script=app.GetCurrentScript()
main_dir=os.path.join(os.getcwd(),'PowerFactory')
pf=PowerFactoryClass.PowerFactoryClass(user_name=user,ShowApp=False)
project2exp = pf.import_project(dict_config[const.strPFProject],main_dir)
project_2exp = pf.activate_project(dict_config[const.strPFProject][:-4])


project2exp.SetAttribute("e:g_file",r"C:\Wentao Zhu\Project\PSA")


export_project = pf.import_project(dict_config[const.strPFProject],main_dir)
export_project.SetAttribute("e:g_file",r"C:\Wentao Zhu\Project\PSA")
location = app.GetCurrentUser()
export_project.g_target=location
export_project.Execute()


importObj=script.CreateObject('CompfdExport','Export')

importObj.SetAttribute("e:g_file",r"C:\Wentao Zhu\Project\PSA")

location=app.GetCurrentUser()
importObj.g_target=location

importObj.Execute()