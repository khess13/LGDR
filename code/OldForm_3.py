'''
LGDR spreadsheet extractor - reads old form and outputs excel with tabs
KHESS 2018-06-08
'''

import pandas as pd
import warnings
import re
import os
import numpy as np
import time
from pprint import pprint as pp


warnings.simplefilter('ignore')

'''UPDATE FILEPATH'''
#filepath for files
inputpath = '/Users/kellyhess/LGDR/'
#filepath for reporting entity names, county, type csv
county_names = pd.read_csv(inputpath+'/ContactInfo.csv')

'''USER INPUT'''
fiscalYear = input('What year? \n')
if len(fiscalYear) < 1: input('Enter a year, dummy. \n')

if fiscalYear != 'test':
	inputpath = inputpath + fiscalYear +'/'


#clean up whitespace around names
county_names['County'] = county_names['County'].apply(lambda x: x.strip())
county_names['ReportingEntityType'] = county_names['ReportingEntityType']\
										.apply(lambda x: x.strip())
county_names['ReportingEntity'] = county_names['ReportingEntity']\
									.apply(lambda x: x.strip())

'''END USER INPUT'''


'''FUNCTIONS'''
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

#moves files aren't found to csvmspath
def move_csv_files(file2move):
	if not os.path.exists(csvmspath):
		os.makedirs(csvmspath)
	#print file2move
	#print trashpath + file2move
	os.rename(inputpath + file2move, csvmspath + file2move)

#find and parse sheet by inputted tab name
def get_tab(tab, filename):
	#try to parse tab
	try:
		xlfile = pd.ExcelFile(inputpath + filename)
		xlfile = xlfile.parse(tab)
		return xlfile
	except:
		print ('Tab ' + tab + ' not in workbook.')
		problemfiles.append(filename +": "+tab)

#use CSV list to derive correct name
def get_right_rep_name(x):
	ct = 0
	for i in repTest:
	    if re.search(i, x['ReportingEntity']):
	        ct += 1
	    #ct += x['ReportingEntity'].count(i)
	return ct

#returns subset of csv for testing
def get_correct_subdf(countyname, repEntityType):
	test_name = county_names[(county_names['County'] == countyname) & \
				(county_names['ReportingEntityType'] == repEntityType)]
	return test_name

#error handling get_correct_subdf returns an empty frame
def error_handle_for_csv(df_to_test, NOF):
	if df_to_test.empty:
		print ('No entries found for '+cnty+' '+repEntType+' '+repEnt)
		cnty = input('Fix county \n') #.encode('utf8')
		repEntType = input('Fix RepEntType \n') #.encode('utf8')
		repEnt = input('Fix RepEnt \n') #.encode('utf8')
		#try call again
		df_to_test = get_correct_subdf(cnty,repEntType)
	if df_to_test.empty:
		print ("Can't find match on CSV")
		move_csv_files(NOF)
		problemfiles.append(NOF)
	return df_to_test

#return the name with highest count
def find_highest_name_csv(df):
	test_name['Count'] = test_name.apply(get_right_rep_name, axis = 1)
	maxValue = df['Count'].agg('max')
	match = df.loc[df['Count'] == maxValue]
	return match

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
csvmspath = inputpath + 'notonlist/'


for filez in filesindir:

	print ('Reading file ', filez)
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
	cnty = raw_excel.iat[0,0].title().strip() #.encode('utf8')
	repEnt = raw_excel.iat[1,0].title().strip() #.encode('utf8')
	fiscYr = pd.to_datetime(raw_excel.iat[3,0], \
			 format = '%d%b%Y').strftime('%m/%d/%Y')
	repEntType = raw_excel.iat[2,0].title().strip() #.encode('utf8')

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
		phoneno.replace('-','').replace('(','').replace(')','') #.encode('utf8')
		faxno.replace('-','').replace('(','').replace(')','') #.encode('utf8')
		#emailadd.encode('utf8')
		#prepBy.encode('utf8')


	#fix for McCormick County name
	cnty = cnty.replace('Mccormick','McCormick')
	repEnt = repEnt.replace('Mccormick','McCormick')
	#Fix school districts with # in name
	repEnt = repEnt.replace('#','')


	#creates repEnt test variable
	repTest = repEnt.split()
	#add regex for complete word match only
	repTest = ['\\'+'b'+x+'\\'+'b' for x in repTest]


	#get sublist for name matching
	test_name = get_correct_subdf(cnty, repEntType)
	#error handle for empty frames
	test_name = error_handle_for_csv(test_name, filez)
	#get name of winner
	good_match = find_highest_name_csv(test_name)

	#error handling for ties
	if len(good_match) > 1:
		anstorepEnt = input('No match, enter name for '+ cnty + ' ' +
															repEntType + ' ' +
															repEnt + '\n')
		print (anstorepEnt)
		yesno = input(' Are you sure? Y/N  \n')

		if yesno == 'N' or yesno == 'n':
			anstorepEnt = input('Okay, try again \n' + repEnt)
		#if no answer given
		if len(anstorepEnt) < 1:
			move_csv_files(filez)
			problemfiles.append(filez)
			failcount += 1
			continue
		#rewrite name of repEnt
		repEnt = anstorepEnt

		#run matching again
		#creates repEnt test variable
		repTest = repEnt.split()
		#add regex for complete word match only
		repTest = ['\\'+'b'+x+'\\'+'b' for x in repTest]

		good_match = find_highest_name_csv(test_name)

		#if no match after reentry, catch for dup counts
		if len(good_match) > 1:
			print (good_match)
			yesno = input('Yes to skip this file?  Y/N')
			if yesno == 'Y' or yesno == 'y':
				print ('No match, moving file')
				move_csv_files(filez)
				failcount += 1
				continue
			repEnt = anstorepEnt
	else:
		#return results
		repEnt = good_match['ReportingEntity'].to_string(index = False)#.encode('utf8')


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
	exportTime = str(time.strftime('%m-%d-%Y'))
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

	print ('Export complete for '+filez)

print ('Problem Files ', pp(problemfiles))
print ('Files analyzed ',filecount)
print ('Files failed ',failcount)
