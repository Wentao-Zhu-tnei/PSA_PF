import argparse as ps
import PSA_Constraint_Resolution as cr

parser=ps.ArgumentParser()
parser.add_argument("filename")
args = parser.parse_args()
print("PSA_PF_auto_CR: " + args.filename)
bOutput2SND = True
bAuto = True
SNDfile = args.filename 
verbose = False
status, msg = cr.PSA_constraint_resolution(bOutput2SND, bAuto, SNDfile)
print("*************************************")
print("Status  = " + str(status))
print("Message = " + msg)
print("*************************************")
