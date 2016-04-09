import os
import csv
import numpy as np
from openpyxl import load_workbook

def writeToCsv(file, data):
	with open(file, 'ab') as file:
		writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
		writer.writerow(data)

def parseTable(a, x1 = None, y1 = None, x2 = None, y2 = None):
	'''
	Remove the csv corresponding to the table if it exists.
	Write the header row then write each data row separately
	'''
	csvFile = "{0}.csv".format(a[2, x1])

	if os.path.isfile(csvFile):
		os.remove(csvFile)

	writeToCsv(csvFile, a[3,x1:x2])
	
	for row in a[y1:y2,x1:x2]:
		writeToCsv(csvFile, row)

def parseWorksheet(a):
	'''
	Parse through the worksheet.
	Determine the row and column ranges.
	Parse the tables once the column range is found.
	'''
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
			parseTable(a, x1, y1, x2, y2)
			x1 = idx + 1
		elif value == 'ET':
			x2 = idx
			parseTable(a, x1, y1, x2, y2)

def main():
	'''
	Load the Excel workbook into a numpy 2D array
	Array is in [rows, columns] format
	Remove the previous csv file
	'''
	wb = load_workbook(filename='roster.xlsx', read_only=True, data_only=True)
	ws = wb['Student']

	parseWorksheet(np.array([[i.value for i in j] for j in ws.rows]))

if __name__ == "__main__":
	main()
