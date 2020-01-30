import os
import csv
import numpy as np
from openpyxl import load_workbook

DATA_DIR = './data'


def write_to_csv(file: str, data: list):
	"""
	Append to a file using the CSV writing format

	Args:
		file: Relative path to the file
		data: Array written to CSV file
	"""
	print(data)
	with open(f'{DATA_DIR}/{file}', 'a') as file:
		writer = csv.writer(
			file,
			delimiter=',',
			quotechar='"',
			quoting=csv.QUOTE_NONNUMERIC
		)
		writer.writerow(data)


def parse_table(a, x1 = None, y1 = None, x2 = None, y2 = None):
	"""
	Remove the csv corresponding to the table if it exists.  Write the header 
	row then write each data row separately.

	Args:
		a: The NumpPy array
		x1: 
		y1:
		x2:
		y2:
	"""
	csv_file = f'{a[2, x1]}.csv'

	if os.path.isfile(csv_file):
		os.remove(csv_file)

	write_to_csv(csv_file, a[3,x1:x2])
	
	for row in a[y1:y2,x1:x2]:
		write_to_csv(csv_file, row)


def parse_worksheet(a):
	"""
	Parse through the worksheet.  Determine the row and column ranges.  Parse 
	the tables once the column range is found.

	Args:
		a: The NumPy array
	"""
	x1 = 0
	x2 = 0
	y1 = 0
	y2 = 0

	for idx,value in enumerate(a[:,0]):
		if value == 'S':
			y1 = idx + 1
		elif value == 'E':
			y2 = idx

	for idx,value in enumerate(a[1,:]):
		if value == 'ST':
			x1 = idx + 1
		elif value == 'BT':
			x2 = idx
			parse_table(a, x1, y1, x2, y2)
			x1 = idx + 1
		elif value == 'ET':
			x2 = idx
			parse_table(a, x1, y1, x2, y2)


def start():
	"""
	Load the Excel workbook into a numpy 2D array.  Array is in [rows, columns] 
	format.  Utilizes OpenPyxl v2.2.5
	"""
	wb = load_workbook(filename='./test/roster.xlsx', read_only=True, data_only=True)
	ws = wb['Student']

	parse_worksheet(np.array([[i.value for i in j] for j in ws.rows]))
