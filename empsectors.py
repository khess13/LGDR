'''
Employment Sectors - takes in BLS data, exports data points for LGDR
KHESS 2018-10-23
'''

import pandas as pd
import time
import re


#FILE LOCATION
'''Home Directory'''
inputpath = '/Users/kellyhess/LGDR/CountyStatData/'
#17 indicates calendar year, 2 indicates quarter
filename = 'allhlcn172.xlsx'
filepath =  inputpath + filename

#WORKBOOK TARGETS
tab = 'US_St_Cn_MSA'
tarcol = 'June Employment'
areatype = 'County'
stname = 'South Carolina'
ownership = ['Private', 'Local Government','State Government',
            'Federal Government']

#DF STUFF
problemfiles = []
curdate = time.strftime('%m/%d/%Y')
exportdate = time.strftime('%m-%d-%Y')
columnorder = ['County','FiscalYear','Sector','Value','DataDate']

#FISCAL YEAR INPUT
fiscYr = raw_input('What fiscal year? MM/DD/YYYY')
if len(fiscYr) < 10:
    fiscYr = '06/30/2018'


#IMPORT FILE
try:
    xlfile = pd.read_excel(filepath, sheet_name = tab)
except:
    print ('Tab ' + tab + ' not in workbook.')
    problemfiles.append(filename +": "+tab)
    exit()

#DATA WRANGLING
#Subset for SC counties
cdtemp = xlfile[(xlfile['Area Type'] == areatype) & \
                (xlfile['St Name'] == stname)]

#Pick up columns for output
cdtemp = cdtemp.iloc[:,[9,10,11,16]] #rows, columns

#new_list = [output FOR iter IN old_list if <condition>]

#filters data for target ownership data, handles subtotals in private
#removes unknown counties
indregex = '^[0-9]{4}\s' #begin with a 4 digit numeric
gardata = 'Unknown'
def filter_ownership(x):
    if re.search(gardata, x['Area']): return False
    elif x['Ownership'] in ownership:
        #remove subtotals from Private Ownership
        if x['Ownership'] == 'Private':
            if re.search(indregex, x['Industry']):
                return True
            else: return False
        return True
    else: return False


yasdata = cdtemp.apply(filter_ownership, axis = 1)
countydata = cdtemp[yasdata]

#remove numbers from Industry text
#regex: beginning line characters repeating 0-9, 1 whitespace
countydata = countydata.replace({'Industry': r'^[0-9]+\s'}, \
                                {'Industry': ''}, regex = True)

#remove County, South Carolina from Area
countydata = countydata.replace({'Area': r'\sCounty, South Carolina'}, \
                                {'Area': ''}, regex = True)

#add fiscal year and export date
countydata['FiscalYear'] = fiscYr
countydata['DataDate'] = curdate


#REFORMAT FOR OUTPUT
#Govt
exportgovt = countydata[(countydata['Ownership'] != 'Private')]\
            .drop(['Industry'], axis = 1)
exportgovt.columns = ['County','Sector','Value','FiscalYear','DataDate']
exportgovt = exportgovt[columnorder]

#Private
exportpriv = countydata[(countydata['Ownership'] == 'Private')]\
            .drop(['Ownership'], axis = 1)
exportpriv.columns = ['County','Sector','Value','FiscalYear','DataDate']
exportpriv = exportpriv[columnorder]

#combine dfs
DFS = [exportgovt, exportpriv]
EmpSectors = pd.concat(DFS)

cntyGroup = EmpSectors.groupby('County')
print (cntyGroup.head())
#for i, r in EmpSectors.iterrows():
'''
#EXPORT TO EXCEL
writer = pd.ExcelWriter(inputpath+'EmpSectors_'+exportdate+'.xlsx')
EmpSectors.to_excel(writer, 'EmploymentSectors', index = False)
writer.save()

print ('Complete!')
'''
