'''
County Stats Upload
KHESS 2018-11-8
'''

import pandas as pd
import time
import re
import os
import sys
import xlsxwriter


#FILE LOCATION
'''Home Directory'''
inputpath = '/Users/kellyhess/LGDR/CountyStatData/'

'''
filenames notes
allhlcn172.xlsx             17 is year, 2 is 2nd quarter
CA1_1969_2016_SC.csv        CA1 is table name
laucnty17.xlsx              17 is year
'''

#File targets
targetfiles = {'empsectors' : 'allhlcn172.xlsx',
               'persincome' : 'CA1_1969_2016_SC.csv',
               'employment' : 'laucnty17.xlsx',
               'population' : 'proj2020.csv'
              }

workbooktargets = { 'persincome' : { 'tab' : targetfiles['persincome'][:-4] + ' - Personal Inc',
                                     'Description' : "Personal income (thousands of dollars)"
                                    },
                    'empsectors' : { 'tab' : 'US_St_Cn_MSA',
                                     'tabcol' : 'June Employment',
                                     'areatype' : 'County',
                                     'stname' : 'South Carolina',
                                     'ownership' : ['Private',
                                                    'Local Government',
                                                    'State Government',
                                                    'Federal Government']
                                    },
                    'population' : { 'tab' : targetfiles['population'][:-4],
                                     'column' : 'July 1, 2018 Projection' #needs to be updated each year!
                                   },
                    'employment' : { 'tab' : targetfiles['employment'][:-5],
                                     'columns' : [ 'County Name/State Abbreviation',
                                                   'Force',
                                                   'Employed',
                                                   'Unemployed',
                                                   '(%)']
                                   }
                }

#Make Ratings & CountySeat tabs
Ratings = pd.DataFrame(columns = ['County','FiscalYear','Moodys',
                                  'Standard&Poors','DataDate'])
CountySeat = pd.DataFrame({'County' :   ['Abbeville',
                                        'Aiken',
                                        'Allendale',
                                        'Anderson',
                                        'Bamberg',
                                        'Barnwell',
                                        'Beaufort',
                                        'Berkeley',
                                        'Calhoun',
                                        'Charleston',
                                        'Cherokee',
                                        'Chester',
                                        'Chesterfield',
                                        'Clarendon',
                                        'Colleton',
                                        'Darlington',
                                        'Dillon',
                                        'Dorchester',
                                        'Edgefield',
                                        'Fairfield',
                                        'Florence',
                                        'Georgetown',
                                        'Greenville',
                                        'Greenwood',
                                        'Hampton',
                                        'Horry',
                                        'Jasper',
                                        'Kershaw',
                                        'Lancaster',
                                        'Laurens',
                                        'Lee',
                                        'Lexington',
                                        'Marion',
                                        'Marlboro',
                                        'Mccormick',
                                        'Newberry',
                                        'Oconee',
                                        'Orangeburg',
                                        'Pickens',
                                        'Richland',
                                        'Saluda',
                                        'Spartanburg',
                                        'Sumter',
                                        'Union',
                                        'Williamsburg',
                                        'York'],
                    'CountySeat' :     ['Abbeville',
                                        'Aiken',
                                        'Allendale',
                                        'Anderson',
                                        'Bamberg',
                                        'Barnwell',
                                        'Beaufort',
                                        'Moncks Corner',
                                        'St. Matthews',
                                        'Charleston',
                                        'Gaffney',
                                        'Chester',
                                        'Chesterfield',
                                        'Manning',
                                        'Walterboro',
                                        'Darlington',
                                        'Dillon',
                                        'St. George',
                                        'Edgefield',
                                        'Winnsboro',
                                        'Florence',
                                        'Georgetown',
                                        'Greenville',
                                        'Greenwood',
                                        'Hampton',
                                        'Conway',
                                        'Ridgeland',
                                        'Camden',
                                        'Lancaster',
                                        'Laurens',
                                        'Bishopville',
                                        'Lexington',
                                        'Marion',
                                        'Bennettsville',
                                        'McCormick',
                                        'Newberry',
                                        'Walhalla',
                                        'Orangeburg',
                                        'Pickens',
                                        'Columbia',
                                        'Saluda',
                                        'Spartanburg',
                                        'Sumter',
                                        'Union',
                                        'Kingstree',
                                        'York']})
CountySeat['DataDate'] = '08/15/2018'

#DF STUFF
problemfiles = []
curdate = time.strftime('%m/%d/%Y')
exportdate = time.strftime('%m-%d-%Y')
columnorder = { 'persincome' : ['County','FiscalYear','PersonalIncome','DataDate'],
                'population' : ['County','FiscalYear','Population','DataDate' ],
                'empsectors' : ['County','FiscalYear','Sector','Value','DataDate'],
                'employment' : ['County','FiscalYear','LaborForce','Employed','Unemployed','Rate','DataDate']
              }
dfs2send = {}

#FISCAL YEAR INPUT
fiscYr = raw_input('What fiscal year? MM/DD/YYYY')
if len(fiscYr) < 10:
    fiscYr = '06/30/2018'

#file error handling ... can't find
def read_file(file2open, tab = ''):
    if 'xls' in file2open:
        try:
            xlfile = pd.read_excel(file2open, sheet_name = tab)
            return xlfile
        except:
            print "Can't read xls file ",file2open
            problemfiles.append(file2open)
    elif 'csv' in file2open:
        try:
            csvfile = pd.read_csv(file2open)
            return csvfile
        except:
            print "Can't read csv file ",file2open
            problemfiles.append(file2open)

#filters ownership data for empsectors
def filter_ownership(x):
    if re.search(gardata, x['Area']): return False
    elif x['Ownership'] in workbooktargets['empsectors']['ownership']:
        #remove subtotals from Private Ownership
        if x['Ownership'] == 'Private':
            if re.search(indregex, x['Industry']):
                return True
            else: return False
        return True
    else: return False

def check_if_files_exist(file2open, expectedfile):
    if not os.path.isfile(file2open):
        print ("Can't find file ",expectedfile,". Check file targets against what's in directory")
        sys.exit()


#IMPORT AND WRANGLE
for target, targetfile in targetfiles.iteritems():
    #print 'Target ', target
    print 'Reading ',targetfile
    #build file path
    filepath = inputpath + targetfile

    #check filenames
    check_if_files_exist(filepath, targetfile)

    #read file
    df = read_file(filepath, workbooktargets[target]['tab'])

    if target == 'persincome':
        #find last column in df
        maxcol = len(df.columns)
        psincm = df.iloc[:,[1,6,(maxcol-1)]].dropna(how = 'any') #rows, column list

        #remove SC totals
        totalfilter = psincm['GeoName'].str.contains('South Carolina state total')
        psincm = psincm[~totalfilter]  #tilde reverses T/F

        #remove all lines except personal Income
        perinfilter = psincm['Description'].str.contains(workbooktargets[target]['Description'], \
                                                        regex = False)
        psincm = psincm[perinfilter]

        #drop Description
        psincm = psincm.drop(['Description'], axis = 1)

        #normalize county names
        psincm = psincm.replace({'GeoName': r', SC'}, \
                        {'GeoName': ''}, regex = True)

        #rename columns
        psincm.columns = ['County', 'PersonalIncome']

        #add additional information
        psincm['FiscalYear'] = fiscYr
        psincm['DataDate'] = curdate

        #reorder columns
        PersonalIncome = psincm[columnorder[target]]

        #append to send
        dfs2send['PersonalIncome'] = PersonalIncome

        #print PersonalIncome.head()

    elif target == 'empsectors':
        #Subset for SC counties
        cdtemp = df[(df['Area Type'] == workbooktargets[target]['areatype']) & \
                    (df['St Name'] == workbooktargets[target]['stname'])]

        #Pick up columns for output
        cdtemp = cdtemp.iloc[:,[9,10,11,16]] #rows, columns

        #filters data for target ownership data, handles subtotals in private
        #removes unknown counties
        indregex = '^[0-9]{4}\s' #begin with a 4 digit numeric
        gardata = 'Unknown'
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
        exportgovt = exportgovt[columnorder[target]]

        #Private
        exportpriv = countydata[(countydata['Ownership'] == 'Private')]\
                    .drop(['Ownership'], axis = 1)
        exportpriv.columns = ['County','Sector','Value','FiscalYear','DataDate']
        exportpriv = exportpriv[columnorder[target]]

        #combine dfs
        DFS = [exportgovt, exportpriv]
        EmpSectors = pd.concat(DFS)

        #append df to send list // [tab name] = df
        dfs2send['EmploymentSectors'] = EmpSectors

        #print EmpSectors.head()

    elif target == 'population':
        #subset population projection, remove NANs
        Population = df.loc[:,['County',workbooktargets[target]['column']]]\
                       .dropna()
        #remove SC total
        Population = Population[Population['County'] != 'South Carolina']

        #rename columns
        Population.columns = ['County', 'Population']

        #add addtional columns
        Population['DataDate'] = curdate
        Population['FiscalYear'] = fiscYr

        #reorder columns
        Population = Population[columnorder[target]]

        #append to send
        dfs2send['Population'] = Population

    elif target == 'employment':
        #get target columns
        labor = df.iloc[5:,[1,3,6,7,8,9]].dropna()

        #get SC data
        labor = labor[labor.iloc[:,0] == '45']

        #drop state code
        labor = labor.drop(['Unnamed: 1'], axis = 1)

        #rename columns
        labor.columns = ['County','LaborForce','Employed','Unemployed',
                              'Rate']

        #normalize county names, removing suffix County, SC
        labor = labor.replace({'County' : r'\sCounty, SC'}, \
                                        {'County' : ''}, regex =  True)

        #add additional info
        labor['FiscalYear'] = fiscYr
        labor['DataDate'] = curdate

        #reorder columns
        Employment = labor[columnorder[target]]

        #append to send
        dfs2send['Employment'] = Employment

        #add holder tabs
        dfs2send['Ratings'] = Ratings
        dfs2send['CountySeat'] = CountySeat

print 'Data wrangling complete!'
print 'Exporting...'
#trim?

#WRITE TO EXCEL
#set tab order for writer
taborder = { 1 : 'Ratings',
             2 : 'PersonalIncome',
             3 : 'Population',
             4 : 'EmploymentSectors',
             5 : 'Employment',
             6 : 'CountySeat'}

#filename of output workbook
writer = pd.ExcelWriter(inputpath+'CountyStatsData_'+exportdate+'.xlsx')

#write dfs to excel file
for key, targ in taborder.iteritems():
    df = dfs2send.get(targ, 'Key not found.')
    print 'Writing ',targ
    df.to_excel(writer, sheet_name = targ, index = False)
'''
for key, df in dfs2send.iteritems():
    df.to_excel(writer, sheet_name = key, index = False)
'''
writer.save()
writer.close()


print 'Problem files: ',problemfiles
print 'Complete!'
