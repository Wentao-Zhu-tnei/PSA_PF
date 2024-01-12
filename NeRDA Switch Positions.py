import os
import re
import pandas as pd
import io
from io import BytesIO
import urllib
from urllib.request import Request, urlopen
from base64 import b64encode
import json
import requests

#Open test excel file and read NeRDA names of circuit breakers
xlsFile = "C:\Wentao Zhu\Project\PSA\Test Data\Test Input Files\zPSA_assetRegFile_4.xlsx"
# xlsFile = 'NeRDA_test_input.xlsx'
tgtCBs = pd.read_excel(xlsFile,sheet_name="Sheet2")
for index, row in tgtCBs.iterrows():
    print(row['NeRDAname'],row['Switch'],row['Status'])

#Open NeRDA API and load circuit breaker data (print first entry in list)
userAndPass = b64encode(b"andrew.maurice@sse.com:GE+FtHXa3eTr2KFUDLBnXmWefCzJdQ8snWxRrPR+5yw=").decode("ascii")
#userAndPass = b64encode(b"transition-technical-trials-psa@sse.com:thcO5484dsURWKDK/wvyJSuupxEwEcJUt26ldIwPRUc=").decode("ascii")
#req = Request(f"https://nerda.opengrid.com/api/all_cbs/")
req = Request(f"https://nerda.opengrid.com/api/transitionstatic/")
#TODO check request error
req.add_header('Authorization', 'Basic %s' % userAndPass)
data= urlopen(req)
JScontent = json.load(data)
CBlist=JScontent[0000]
with open('data.json', 'w') as f:
    json.dump(JScontent, f)

print(CBlist['switches'][0]['switch_name'], CBlist['switches'][0]['nerda_switch_uuid'], CBlist['switches'][0]['measurements'])

for cb in CBlist:
    print(cb['Switches'][0])
    #print(cb['shortName'],cb['shortName'][:9],cb['value'])
    #print(cb['Switches'][0]['switch_name'], cb['Switches'][0]['nerda_switch_uuid'], cb['Switches'][0]['measurements'])

for index, row in tgtCBs.iterrows():
    #tgtCBs.at[index,'Status']='Not found'
    for cb in CBlist:
        if row['NeRDAname'][:9] == cb['shortName'][:9]:
            print(cb['shortName'],cb['value'],row['NeRDAname'])
            tgtCBs.at[index,'Switch'] = cb['value']
            if cb['value']==0:
                tgtCBs.at[index,'Status']='Closed'
            elif cb['value']==1:
                tgtCBs.at[index,'Status']='Open'
            else:
                tgtCBs.at[index,'Status']='Undefined (' + str(cb['value']) + ')'
                tgtCBs.at[index,'Switch'] = -1

for index, row in tgtCBs.iterrows():
    print(row['NeRDAname'], row['Switch'], row['Status'])

tgtCBs.to_excel('NeRDA_test_output.xlsx', index = False)
