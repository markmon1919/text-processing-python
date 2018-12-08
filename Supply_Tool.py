import os
import csv
import pandas as pd
from tkinter import *
from tkinter import messagebox

def rename_xls():
	#Renaming file extension case
	for files in os.listdir(pwd):
		if files.endswith('.XLS'):
			os.rename(files, files.replace('.XLS', '.xls'))
		if files.endswith('.XLSX'):
			os.rename(files, files.replace('.XLSX', '.xlsx'))

def rename_space():
	print('Removing whitespaces...')
	for files in os.listdir(pwd):
		if files.endswith('.xls') or files.endswith('.xlsx') or files.endswith('.csv'):
			os.rename(files, files.replace(' ', '_'))

def get_xls():
	for files in os.listdir(pwd):
		if files.endswith('.xls') or files.endswith('.xlsx'):
			xlsLs.append(files)
			print('  [*]', files)
	print(' ', len(xlsLs), 'Excel files found...\n')

def convert():
	print('Converting excel files to CSV UTF-8 format...')
	file1_xls = pd.read_excel(supplyXLS, 'Sheet1', index_col=None)
	file1_xls.to_csv(supplyXLS.replace('.xlsx', '') + '.csv', sep=',', encoding='UTF-8')

	file2_xls = pd.read_excel(hcXLS, 'Sheet1', index_col=None)
	file2_xls.to_csv(hcXLS.replace('.xlsx', '') + '.csv', sep=',', encoding='UTF-8')

	# file3_xls = pd.read_excel(forXLS, 'Conso', index_col=None)
	# file3_xls.to_csv(forXLS.replace('.xlsx', '') + '.csv', sep='\t', encoding='UTF-8')

def get_csv():
	for files in os.listdir(pwd):
		if files.endswith('.csv'):
			csvLs.append(files)
			print('  [*]', files)
	print(' ', len(csvLs), 'CSV files found...\n')

def del_col():
	print('Deleting columns from Supply_To_Be_Deleted...\n')
	with open(forCSV, 'r', encoding='UTF-8') as for_csv:
		df = pd.read_csv(for_csv, low_memory=False)
		with open(supplyCSV, 'r') as supply_csv:
			df2 = pd.read_csv(supply_csv)
			for i in df2.columns:
				try:
					output = df.drop([i], axis=1)	
				except KeyError:
					pass
		output.to_csv('draft.csv', index=False)

def filter_rows():
	#latest sheet(#FILTER 1,511 records)
	print('Filtering Rows...\n')
	with open('draft.csv', 'r', encoding='UTF-8') as draft_csv:
		df = pd.read_csv(draft_csv, low_memory=False)
		df = df.loc[df['IG'].isin(['SFDC IPS', 'Oracle IPS', 'Workday IPS']) | df['Resources Reqd From'].isin(['Salesforce IPS', 'Oracle IPS', 'Workday IPS'])]
		output = df.drop('Technology', axis=1)
		output.to_csv('draft.csv', index=False)

def vlookup():
	print('VLOOKUP', hcCSV + str('&') + forCSV, '...')
	with open('draft.csv', 'r', encoding='UTF-8') as draft_csv:
		df = pd.read_csv(draft_csv)
		with open(hcCSV, 'r', encoding='UTF=8') as hc_csv:
			df2 = pd.read_csv(hc_csv)
			df2 = df2[['Name', 'Technology']]
			output = df.merge(df2, on=['Name'], how='outer')
			output.to_csv('output.csv', index=False)
	#take not null values or drop null values and 
	#replace all NULL to Other in Technology Column
	with open('output.csv', 'r', encoding='UTF-8') as output_csv:
		df = pd.read_csv(output_csv)
		df['Technology'] = df['Technology'].fillna('Other')
		output = df[pd.notnull(df['Personnel No'])]
		output.to_csv('output.csv', index=False)

def save_output():
	os.remove('draft.csv')
	os.remove(supplyCSV)
	os.remove(hcCSV)
	#os.remove(forCSV)

	fn = input('Enter output filename: ')
	print('Saving output file as: \n', ' [*]', fn + str('.csv'))
	try:
		os.rename('output.csv', fn + '.csv')
	except FileExistsError:
		os.remove(fn + '.csv')
		os.rename('output.csv', fn + '.csv')

	root = Tk()
	root.withdraw()
	messagebox.showinfo(title='NOTE: ', message='\nPlease find and replace all characters "Ã±" to "ñ" manually...\nClick OK to open the output file.')

	print('Opening output file: \n', ' [*]', pwd + str('\\') + fn + str('.csv'))
	os.startfile(fn + '.csv')

if __name__ == '__main__':

	print('\nSupply Automation Tool')
	print('Created by Mark Mon Monteros')
	print('Python ver 3.7*\n')
	print('NOTE: Please convert manually password-protected files before executing this program.\n')
	
	pwd = os.path.dirname(os.path.realpath(__file__))
	xlsLs = []
	csvLs = []

	rename_xls()
	rename_space()
	get_xls()
	#Assign XLS Sheets
	for sheet in xlsLs:
		if 'Supply_To_Be_Deleted' in sheet:
			supplyXLS = sheet
		elif 'HC' in sheet:
			hcXLS = sheet
		elif 'for_reporting' in sheet:
			forXLS = sheet

	convert()
	get_csv()
	#Assign CSV Sheets
	for sheet in csvLs:
		if 'Supply_To_Be_Deleted' in sheet:
			supplyCSV = sheet
		elif 'HC' in sheet:
			hcCSV = sheet
		elif 'for_reporting' in sheet:
			forCSV = sheet

	del_col()
	filter_rows()
	vlookup()
	save_output()

	print('\nDONE!')
