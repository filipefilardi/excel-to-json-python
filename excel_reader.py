#!/usr/bin/python
# -*- coding: utf-8 -*-	

import os.path
import json

from openpyxl import Workbook
from openpyxl import load_workbook

def get_titles(worksheet):
	titles = []
	for column in range(worksheet.max_column):
		try:
			titles.append(worksheet.cell(row=1, column=column+1).value.upper())
		except:
			titles.append(worksheet.cell(row=1, column=column+1).value)
	return titles


def all_data_to_json(worksheet):
	#data = []
	#titles = get_titles(worksheet)

	with open('teste.json', 'w') as file:
		json_data = []

		for row in range(1, worksheet.max_row):
			item = {}
			for column in range(worksheet.max_column):
				# data.append(worksheet.cell(row=row+1, column=column+1).value) 
				try:
					item[worksheet.cell(row=1, column=column+1).value.upper()] = worksheet.cell(row=row+1, column=column+1).value.encode('utf-8')
				except:
					try:
						item[worksheet.cell(row=1, column=column+1).value.upper()] = worksheet.cell(row=row+1, column=column+1).value
					except:
						item[worksheet.cell(row=1, column=column+1).value] = worksheet.cell(row=row+1, column=column+1).value
			json_data.append(item)

		json.dump(json_data, file, indent = 4, ensure_ascii = False) # sort_keys = True
		file.close()

def main():
	wb = load_workbook(filename='teste1.xlsx' , read_only=True)
	sheets = wb.get_sheet_names()

	for sheet in sheets:
		ws = wb[sheet]

		all_data_to_json(ws)

if __name__ == '__main__':
	main()