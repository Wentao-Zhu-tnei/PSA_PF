import argparse as ps
import PSA_PF_SF as sf

parser=ps.ArgumentParser()
parser.add_argument("filename")
args = parser.parse_args()
print("PSA_PF_auto_SF: " + args.filename)
bOutput2SND = True
bAuto = True
SND_file = args.filename 
verbose = False
status, msg = sf.PSA_PF_workflowSF(bOutput2SND, bAuto, SND_file, verbose=False)
print("*************************************")
print("Status  = " + str(status))
print("Message = " + msg)
print("*************************************")
