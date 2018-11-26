import pandas as pd
import os

inputpath = '/Users/kellyhess/LGDR/CountyStatData/allhlcn172.xlsx'
#filepath = '/Users/kellyhess/LGDR/'
problemfiles = []

def read_file(file2open):
    if 'xls' in file2open:
        try:
            xlfile = pd.read_excel(file2open, sheet_name = 'US_St_Cn_MSA')
            return xlfile
        except:
            print "Can't read xls file " + file2open
            problemfiles.append(file2open)
    elif 'csv' in file2open:
        try:
            csvfile = pd.read_csv(file2open)
            return csvfile
        except:
            print "Can't read csv file " + file2open
            problemfiles.append(file2open)



xlfile = read_file(inputpath)
print xlfile

#print wbtabs
