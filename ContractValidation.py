from openpyxl import workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.formula import Tokenizer
import pandas as pd
import time


'''

read_file_Mac = pd.read_csv (r'C:\\Users\\Gert\\Documents\\Development\\excellekes\\final\\MAC.csv', on_bad_lines='skip')
read_file_Mac.to_excel (r'C:\\Users\\Gert\\Documents\\Development\\excellekes\\final\\MAC.xlsx', index = None, header=True)

read_file_Att = pd.read_csv (r'C:\\Users\\Gert\\Documents\\Development\\excellekes\\final\\ATT.csv', on_bad_lines='skip')
read_file_Att.to_excel (r'C:\\Users\\Gert\\Documents\\Development\\excellekes\\final\\ATT.xlsx', index = None, header=True)
'''




path1 = 'C:\\Users\\Gert\\Documents\\Development\\excellekes\\final\\MAC.xlsx'
path2 = 'C:\\Users\\Gert\\Documents\\Development\\excellekes\\final\\ATT.xlsx'
pathdest = 'C:\\Users\\Gert\\Documents\\Development\\excellekes\\final\\RentalContract.xlsx'


def addToLastColumn(path, addition):
	wb = load_workbook(filename=path)
	wbs0 = wb.worksheets[0]
	#print(wbs0.max_column)
	char = get_column_letter(wbs0.max_column+1)
	#print(char)
	i = 1
	for row in wbs0:
		wbs0[char + str(i)] = addition
		i = i+1
	wb.save(path) 
	print('Added ' + addition + ' ' + str(wbs0.max_row) + ' times')


def vlookupToLastCollumn(path, title, sheet, formula):
	wb = load_workbook(filename=path)
	wbs0 = wb.worksheets[0]
	#print(wbs0.max_column)
	char = get_column_letter(wbs0.max_column+1)
	#print(char)
	i = 1
	for row in wbs0:
		#wbs0[char + str(i)] = "=VLOOKUP(AB2,'Machines + Attachments'!B:S,10,FALSE)"
		wbs0[char + str(i)] = "=VLOOKUP(AB" + str(i) + ",'" + sheet + formula
		i = i+1
	wbs0[char + '1'] = title
	wb.save(path) 
	
#VLOOKUP(AB2,'Machines + Attachments'!B:S,10,FALSE)"

#Purpose = "'!B:L;10,FALSE)"

def sheetCopyPaste(ws, destination):
	for row in ws:
		for cell in row:
			destination[cell.coordinate].value = cell.value



def rowMemory(ws):  #produce the list of items in the particular row
        for row in ws.iter_rows(min_row=2):
            yield [cell.value for cell in row]


def combiningWorkbooks(title, destination, header, second):

	wbh = load_workbook(filename=header)
	wsh_0 = wbh.worksheets[0]

	wb2 = load_workbook(filename=second)
	ws2_0 = wb2.worksheets[0]

	'''
	wb3 = load_workbook(filename=third, read_only=True)
	ws3_0 = wb3.worksheets[0]
	ws3_0.delete_cols(1)

	wb4 = load_workbook(filename=fourth, read_only=True)
	ws4_0 = wb4.worksheets[0]
	ws4_0.delete_cols(1)
	'''
	wbdestination = load_workbook(filename=destination)
	wsdestination_1 = wbdestination.create_sheet(title)

	
	sheetCopyPaste(wsh_0, wsdestination_1)

	list_to_append = list(rowMemory(ws2_0))
	#print(list_to_append)
	for items in list_to_append:
		#print(items)
		wsdestination_1.append(items)
			

	wbdestination.save(destination)
	wbh.close()
	wb2.close()


machAttsheetTitle = 'Machines + Attachments'

	
# 	validation starters
Purpose = "'!B:L,10,FALSE)"

#print('Start Preparing Machine and Attachment data')
#time.sleep(1)
#addToLastColumn(path1, 'MACHINE')
#addToLastColumn(path2, 'ATTACHMENT')
#time.sleep(1)
'''
wbh = load_workbook(filename=path1)
wsh_0 = wbh.worksheets[0]
wbdestination = load_workbook(filename=pathdest)
wsdestination_1 = wbdestination.create_sheet("title")

sheetCopyPaste(wsh_0, wsdestination_1)
wbdestination.save(pathdest)
'''

#combiningWorkbooks(machAttsheetTitle, pathdest, path1, path2)
'''
wb = load_workbook(filename=pathdest)
wbs0 = wb.worksheets[0]
wbs0['BJ2'] = "=VLOOKUP(AB2,'Machines + Attachments'!B:S,10,FALSE)"
wb.save(pathdest)
'''
#wbs0['BJ2'] = Tokenizer("""=VERT.ZOEKEN(AB1;'Machines + Attachments'!B:L;10)""")

vlookupToLastCollumn(pathdest, 'Purpose', machAttsheetTitle, Purpose)

