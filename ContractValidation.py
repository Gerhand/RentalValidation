from openpyxl import workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.formula import Tokenizer
import pandas as pd
import time
import ValidationFunctions as vf


'''

read_file_Mac = pd.read_csv (r'C:\\Users\\Gert\\Documents\\Development\\excellekes\\final\\MAC.csv', on_bad_lines='skip')
read_file_Mac.to_excel (r'C:\\Users\\Gert\\Documents\\Development\\excellekes\\final\\MAC.xlsx', index = None, header=True)

read_file_Att = pd.read_csv (r'C:\\Users\\Gert\\Documents\\Development\\excellekes\\final\\ATT.csv', on_bad_lines='skip')
read_file_Att.to_excel (r'C:\\Users\\Gert\\Documents\\Development\\excellekes\\final\\ATT.xlsx', index = None, header=True)
'''



path1 = 'FileBucket\\MAC.xlsx'
path2 = 'FileBucket\\ATT.xlsx'
pathdest = 'FileBucket\\RentalContract.xlsx'
pathtest = 'FileBucket\\test.xlsx'


machAttsheetTitle = 'Machines + Attachments'

	
# 	validation starters
Purpose = "'!B:L,10,FALSE)"
State = "'!B:L,11)"
ObjectType = "'!B:AH,33)"
InFleetStartDate = ""
InFleetEndDate = "'!B:L,9)"

valcol_a = 'AddressLine2'
valcol_b = 'ZipCode'
valcol_c = 'City'
valcol_d = 'CountryISO2'
#valcol_e = 'Latitude'
#valcol_f = 'Longitude'
#valcol_g = 'EstimatedEndDate'
#valcol_h = 'DeliveryEarliest'
#valcol_h = 'DeliveryLatest'
valcol_h = 'Purpose'


ValidationRegime = int(input('What kind of validation do you want? 1 = Full, 2 = Combining and validation, 3 = File Validation only\n'))

if ValidationRegime == 1:
	print('Start Preparing Machine and Attachment data', flush=True)
	time.sleep(1)
	vf.addToLastColumn(path1, 'MACHINE')
	vf.addToLastColumn(path2, 'ATTACHMENT')
	time.sleep(1)

	vf.combiningWorkbooks(machAttsheetTitle, pathdest, path1, path2)


	wb = load_workbook(filename=pathdest)
	wbs0 = wb.worksheets[0]

	wbs0.auto_filter.ref = wbs0.dimensions
	wbs0.freeze_panes = 'A2' 

	print('Comparing data', flush=True)
	vf.vlookupToLastCollumn(wbs0, 'Purpose', machAttsheetTitle, Purpose)
	vf.vlookupToLastCollumn(wbs0, 'State', machAttsheetTitle, State)
	vf.vlookupToLastCollumn(wbs0, 'ObjectType', machAttsheetTitle, ObjectType)
	vf.vlookupToLastCollumn(wbs0, 'InFleetEndDate', machAttsheetTitle, InFleetEndDate)

	wb.save(pathdest)


elif ValidationRegime == 2:
	vf.combiningWorkbooks(machAttsheetTitle, pathdest, path1, path2)


	wb = load_workbook(filename=pathdest)
	wbs0 = wb.worksheets[0]

	wbs0.auto_filter.ref = wbs0.dimensions
	wbs0.freeze_panes = 'A2' 

	print('Comparing data', flush=True)
	vf.vlookupToLastCollumn(wbs0, 'Purpose', machAttsheetTitle, Purpose)
	vf.vlookupToLastCollumn(wbs0, 'State', machAttsheetTitle, State)
	vf.vlookupToLastCollumn(wbs0, 'ObjectType', machAttsheetTitle, ObjectType)
	vf.vlookupToLastCollumn(wbs0, 'InFleetEndDate', machAttsheetTitle, InFleetEndDate)
	wb.save(pathdest)


	print('Validating', flush=True)

	vf.searchForBlanks(wb, wbs0, valcol_a)
	vf.searchForBlanks(wb, wbs0, valcol_b)
	vf.searchForBlanks(wb, wbs0, valcol_c)
	vf.searchForBlanks(wb, wbs0, valcol_d)
	vf.searchForBlanks(wb, wbs0, valcol_e)

	wb.save(pathdest)


elif ValidationRegime == 3:
	wb = load_workbook(filename=pathdest)
	wbs0 = wb.worksheets[0]

	wbs0.auto_filter.ref = wbs0.dimensions
	wbs0.freeze_panes = 'A2' 	

	print('Comparing data', flush=True)

	vf.vlookupToLastCollumn(wbs0, 'Purpose', machAttsheetTitle, Purpose)
	vf.vlookupToLastCollumn(wbs0, 'State', machAttsheetTitle, State)
	vf.vlookupToLastCollumn(wbs0, 'ObjectType', machAttsheetTitle, ObjectType)
	vf.vlookupToLastCollumn(wbs0, 'InFleetEndDate', machAttsheetTitle, InFleetEndDate)

	print('Validating', flush=True)

	vf.searchForBlanks(wb, wbs0, valcol_a)
	vf.searchForBlanks(wb, wbs0, valcol_b)
	vf.searchForBlanks(wb, wbs0, valcol_c)
	vf.searchForBlanks(wb, wbs0, valcol_d)
	vf.searchForBlanks(wb, wbs0, valcol_e)





	wb.save(pathdest)





	'''




wb = load_workbook(filename=pathtest)
wbs0 = wb.worksheets[0]
vf.searchForBlanks(wb, wbs0, 'borst')

#print(wbs0['C4'].value)

wb.save(filename=pathtest) '''
