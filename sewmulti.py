'''
LGDR spreadsheet extractor
KHESS 2018-09-18
'''

import pandas as pd
import warnings
import re
import os
import numpy as np
import time
import multiprocessing



'''USER INPUT'''
fiscalYear = raw_input('What year?')
if len(fiscalYear) < 1: raw_input('Enter a year, dummy.')

inputpath = '/Users/kellyhess/LGDR/'

if fiscalYear != 'test':
	inputpath = inputpath + fiscalYear +'/'
'''END USER INPUT'''


'''UTIL FUNCTIONS'''
from os import listdir

#Find XLS* files in directory
def find_xlsx_files( path_to_dir, suffix='.xls'):
    filenames = listdir(path_to_dir)
    #return [ filename for filename in filenames if filename.endswith( suffix ) ]
    return [filename for filename in filenames if suffix in filename]
    #selected item from list FOR iter IN list IF <condition>

#Find target sheet in workbook
def find_tab(tabtab,xlsxf):
	for t in xlsxf:
		if not re.search(tabtab,t): continue
		getsheet = t
		#print 'Found ', getsheet
		return getsheet

#moves files that won't open to trashfiles
def move_trash_files(file2move):
	if not os.path.exists(trashpath):
		os.makedirs(trashpath)
	#print file2move
	#print trashpath + file2move
	os.rename(inputpath + file2move, trashpath + file2move)


#find and parse sheet by inputted tab name
def get_tab(tab, filename):
	#try to parse tab
	try:
		xlfile = pd.ExcelFile(inputpath + filename)
		xlfile = xlfile.parse(tab)
		return xlfile
	except:
		print 'Tab ' + tab + ' not in workbook.'
		problemfiles.append(filename +": "+tab)

'''END FUNCTIONS'''


'''TARGETS'''
#workbook tab names
wstarg = ['General Data','General Obligation','Revenue',
		  'County Supplemental Data']


'''END TARGETS'''

'''SETUP DFs/Exception Lists'''
problemfiles = []

#Debt column names
#colOrder = ['Purpose','BeginFY','IssuedFY','RetiredFY','EndFY']
colOrder = ['Tab','County','ReportingEntity','ReportingEntityType',
			'FiscalYear','DebtType','Term','Purpose','BeginFY','IssuedFY',
			'RetiredFY','DueNFY','EndFY']

colRename = ['Purpose','BeginFY','IssuedFY','RetiredFY','DueNFY','EndFY']

'''END DFs'''

#gen file list
filesindir = find_xlsx_files(inputpath)
#filecount = len(filesindir)
filecount = 0
failcount = 0
trashpath = inputpath+'trashfiles/'


for filez in filesindir:

	print 'Reading file ', filez
	filecount += 1

	'''General Data----ContactInfo'''
	#parse sheet
	xlfile = get_tab('General Data', filez)
	#catch problems with tab retrieval
	if xlfile is None:
		move_trash_files(filez)
		failcount += 1
		continue

	raw_excel = xlfile.iloc[1:8,1:5]  #rows, columns

	#Repeating variables
	cnty = raw_excel.iat[0,0].encode('utf8')
	repEnt = raw_excel.iat[1,0].encode('utf8')
	fiscYr = pd.to_datetime(raw_excel.iat[3,0], \
			 format = '%d%b%Y').strftime('%m/%d/%Y')
	repEntType = raw_excel.iat[2,0].encode('utf8')

	phoneno = raw_excel.iat[6,0]
	faxno = raw_excel.iat[6,2]
	emailadd = raw_excel.iat[5,0]
	prepBy = raw_excel.iat[4,0]

	#check for numerical value on all fields
	if isinstance(phoneno,(int,float))  or isinstance(faxno,(int,float)) \
		or isinstance(emailadd,(float,int)) or isinstance(prepBy,(int,float)):
		phoneno = str(phoneno).replace('-','').replace('(','').replace(')','')
		faxno = str(faxno).replace('-','').replace('(','').replace(')','')
		emailadd = str(emailadd)
		prepBy = str(prepBy)
	else:
		phoneno.encode('utf8').replace('-','').replace('(','').replace(')','')
		faxno.encode('utf8').replace('-','').replace('(','').replace(')','')
		emailadd.encode('utf8')
		prepBy.encode('utf8')

	# .encode('utf8') to remove u'
	buildFrame = {'County': [cnty],
				'ReportingEntity': [repEnt],
				'ReportingEntityType': [repEntType],
				'FiscalYear': [fiscYr],
				'PreparedBy': [prepBy],
				'Email': [emailadd],
				'Phone': [phoneno],
				'Fax': [faxno]
				}


	gendata = pd.DataFrame(data=buildFrame).fillna(0) #.reset_index()

	#reorder columns
	ContactInfo = gendata[['County','ReportingEntity','ReportingEntityType',
						'FiscalYear','PreparedBy','Email','Phone','Fax']]



	'''General Obligation & Revenue---Debt'''
	#parse sheets
	#GENERAL OBLIGATION
	xlfile = get_tab('General Obligation', filez)

	genOb = xlfile.iloc[6:16,0:6].fillna(0)
	genOb.columns = colRename

	othgenOb = xlfile.iloc[18:24,0:6].fillna(0)
	othgenOb.columns = colRename

	genObS = xlfile.iloc[29:39,0:6].fillna(0)
	genObS.columns = colRename

	othgenObS = xlfile.iloc[41:47, 0:6].fillna(0)
	othgenObS.columns = colRename

	#REVENUE
	xlfile1 = get_tab('Revenue', filez)
	genObRev = xlfile1.iloc[6:16,0:6].fillna(0)
	genObRev.columns = colRename

	othgenObRev = xlfile1.iloc[18:24,0:6].fillna(0)
	othgenObRev.columns = colRename

	genObRevS = xlfile1.iloc[29:39,0:6].fillna(0)
	genObRevS.columns = colRename

	othgenObRevS = xlfile1.iloc[41:47,0:6].fillna(0)
	othgenObRevS.columns = colRename


	#add additional identifiers
	#GENERAL OBLIGATION
	genOb['Tab'] = 'General Obligation'
	genOb['County'] = cnty
	genOb['ReportingEntity'] = repEnt
	genOb['ReportingEntityType'] = repEntType
	genOb['FiscalYear'] = fiscYr
	genOb['DebtType'] = 'General Obligation'
	genOb['Term'] = 'Long Term'

	genObS['Tab'] = 'General Obligation'
	genObS['County'] = cnty
	genObS['ReportingEntity'] = repEnt
	genObS['ReportingEntityType'] = repEntType
	genObS['FiscalYear'] = fiscYr
	genObS['DebtType'] = 'General Obligation'
	genObS['Term'] = 'Short Term'

	#OTHER GENERAL
	othgenOb['Tab'] = 'General Obligation'
	othgenOb['County'] = cnty
	othgenOb['ReportingEntity'] = repEnt
	othgenOb['ReportingEntityType'] = repEntType
	othgenOb['FiscalYear'] = fiscYr
	othgenOb['DebtType'] = 'Other General Obligation'
	othgenOb['Term'] = 'Long Term'

	othgenObS['Tab'] = 'General Obligation'
	othgenObS['County'] = cnty
	othgenObS['ReportingEntity'] = repEnt
	othgenObS['ReportingEntityType'] = repEntType
	othgenObS['FiscalYear'] = fiscYr
	othgenObS['DebtType'] = 'Other General Obligation'
	othgenObS['Term'] = 'Short Term'

	#REVENUE
	genObRev['Tab'] = 'Revenue'
	genObRev['County'] = cnty
	genObRev['ReportingEntity'] = repEnt
	genObRev['ReportingEntityType'] = repEntType
	genObRev['FiscalYear'] = fiscYr
	genObRev['DebtType'] = 'Revenue'
	genObRev['Term'] = 'Long Term'

	genObRevS['Tab'] = 'Revenue'
	genObRevS['County'] = cnty
	genObRevS['ReportingEntity'] = repEnt
	genObRevS['ReportingEntityType'] = repEntType
	genObRevS['FiscalYear'] = fiscYr
	genObRevS['DebtType'] = 'Revenue'
	genObRevS['Term'] = 'Short Term'

	#OTHER REVENUE
	othgenObRev['Tab'] = 'Revenue'
	othgenObRev['County'] = cnty
	othgenObRev['ReportingEntity'] = repEnt
	othgenObRev['ReportingEntityType'] = repEntType
	othgenObRev['FiscalYear'] = fiscYr
	othgenObRev['DebtType'] = 'Other Revenue'
	othgenObRev['Term'] = 'Long Term'

	othgenObRevS['Tab'] = 'Revenue'
	othgenObRevS['County'] = cnty
	othgenObRevS['ReportingEntity'] = repEnt
	othgenObRevS['ReportingEntityType'] = repEntType
	othgenObRevS['FiscalYear'] = fiscYr
	othgenObRevS['DebtType'] = 'Other Revenue'
	othgenObRevS['Term'] = 'Short Term'


	#reorder columns
	genOb = genOb[colOrder]
	genObS = genObS[colOrder]
	othgenOb = othgenOb[colOrder]
	othgenObS = othgenObS[colOrder]
	genObRev = genObRev[colOrder]
	genObRevS = genObRevS[colOrder]
	othgenObRev = othgenObRev[colOrder]
	othgenObRevS = othgenObRevS[colOrder]

	#combine into one DF
	DFS = [genOb,genObS,othgenOb,othgenObS,genObRev,othgenObRev,
			genObRevS,othgenObRevS]
	Debt = pd.concat(DFS)

	#Drop DueNFY -- discontinued in next form
	Debt = Debt.drop(Debt.columns[11], axis = 1) #was 10

	#print Debt

	'''County Supplemental Data---CountyStatistics'''
	#TAX DATA

	#parse sheet
	xlfile = get_tab('County Supplemental Data', filez)
	TaxData = xlfile.iloc[7:20,0:4].fillna(0)

	#column 1 doesn't contain data
	TaxData = TaxData.drop(TaxData.columns[1], axis = 1)

	#update column names
	TaxData.columns = ['Statistic','StatisticValue','StatisticPercent']


	#drop section headers / garbage rows
	TaxData = TaxData.drop([9])
	TaxData = TaxData.drop([13])


	#add additional identifiers
	TaxData['Tab'] = 'County Supplemental Data'
	TaxData['County'] = cnty
	TaxData['ReportingEntity'] = repEnt
	TaxData['FiscalYear'] = fiscYr
	TaxData['Section'] = ''
	TaxData['Category'] = ''

	#conditional additional identifiers
	for index, row in TaxData.iterrows():
		if row['Statistic'] in ['8% of Assessed Property Valuation', 'Total General Obligation Debt Outstanding','Debt Margin']:
			TaxData.loc[index, 'Category'] = 'Debt Limit'
			TaxData.loc[index, 'Section'] = 'Tax Data'
		elif row['Statistic'] in ['Assessed Property Valuation','Current Tax Collections']:
			TaxData.loc[index, 'Category'] = 'General Tax Data'
			TaxData.loc[index, 'Section'] = 'Tax Data'
		elif row['Statistic'] in ['Property Taxes','State Aid','Federal Aid','Fees, Fines and Forfeitures','Interest Income','Other']:
			TaxData.loc[index, 'Category'] = 'Revenue Sources'
			TaxData.loc[index, 'Section'] = 'Tax Data'

	#reorder columns
	TaxData = TaxData[['Tab','County','ReportingEntity','FiscalYear',
	'Section','Category','Statistic','StatisticValue','StatisticPercent']]

	#print TaxData

	#ECONOMIC PROFILE
	#parse sheet
	EconProf = xlfile.iloc[23:28,0:4].fillna(0)

	#drop 1 empty column
	EconProf = EconProf.drop(EconProf.columns[1:2], axis = 1)

	#rename columns
	EconProf.columns = ['Statistic','StatisticPercent','StatisticValue']

	#add additional identifiers
	EconProf['Tab'] = 'County Supplemental Data'
	EconProf['County'] = cnty
	EconProf['ReportingEntity'] = repEnt
	EconProf['FiscalYear'] = fiscYr
	EconProf['Section'] = 'Economic Profile'
	EconProf['Category'] = 'Major Employers'

	#reorder columns
	EconProf = EconProf[['Tab','County','ReportingEntity','FiscalYear',
	'Section','Category','Statistic','StatisticValue','StatisticPercent']]

	#print EconProf


	#combine to make CountyStatistics
	DFS1 = [TaxData,EconProf]
	CountyStatistics = pd.concat(DFS1).fillna(0)

	#print CountyStatistics

	#print CountyStatistics
	CountyStatistics = CountyStatistics[['Tab','County','ReportingEntity',
										'FiscalYear','Section','Category',
										'Statistic','StatisticValue',
										'StatisticPercent']]
	#print CountyStatistics


	'''Write to Excel'''

	#Make Export Spreadsheet
	exportTime = time.strftime('%m-%d-%Y')
	outputpath = inputpath + 'output/'
	fiscYr = fiscYr.replace('/','-') #remove / so not read as file dir

	#check if folder exists
	if not os.path.exists(outputpath):
		os.makedirs(outputpath)

	#LGDR_County_EntityType_EntityName_FiscalYear_DateStamp
	writer = pd.ExcelWriter(outputpath+'LGDR_'+cnty+'_'+
							repEnt+'_'+repEntType+'_'+fiscYr+'_'+
							exportTime+'.xlsx')


	#write sheets
	Debt.to_excel(writer, 'Debt', index = False)
	CountyStatistics.to_excel(writer, 'CountyStatistics', index = False)
	ContactInfo.to_excel(writer, 'ContactInfo', index = False)

	writer.save()

	print 'Export complete for '+filez

print problemfiles
print 'Files analyzed '+str(filecount)
print 'Files failed '+str(failcount)
