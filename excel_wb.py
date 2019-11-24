# Homework
# Given a file directory,
# Find all excel workbooks,
# Return a new csv file with each row being a sum of
# all values in that row of all sheets of all workbooks

import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import os
import csv
import numpy as np

# Define global variable
workbooks = []

# Find the Excel workbooks in a file directory and add them to workbooks list. 
# Go through each level of the directory 
# checking for files with extensions ".xlsx" or ".xls"
def find_workbooks(fdir_path):
	
	# Create list (of strings) of all files and directories in the given directory
	filelist = os.listdir(fdir_path)

	# Check whether each item is a workbook or directory
	for f in filelist:
		
		# Convert file name to lowercase
		fname = f.lower()
		
		# If file name ends in Excel workbook extenstions, add its path to list of workbooks
		if fname[-4:]=='.xls' or fname[-5:]=='.xlsx':
			fpath = fdir_path+'/'+f
			workbooks.append(fpath)	# file added to workbooks		

		# Else if file is a directory, store the file path 
		# and find workbooks in the subdirectory
		elif os.path.isdir(fdir_path+'/'+f):
			path = fdir_path+'/'+f		
			find_workbooks(path) # now looking through subdirectory


def sum_sheet_rows(df):
	index = df.index
	columns = df.columns
	values = df.values

	num_rows = len(index)
	num_columns = len(columns)

	row_sums = [0]*num_rows

	for r in range(num_rows):
		row_sum = 0
		for c in range(num_columns):
			cell = values[r][c]
			#print(str(cell) + ' -- ' + str(type(cell)) )
			
			# check for numpy.float64 and float values (could be nan or number)
			if type(cell) == np.float64 or type(cell) == float :
				# check if cell is not a nan value
				if not np.isnan(cell):
					# add to row count
					row_sum+=cell 
			
			# check for int values
			elif type(cell) == int:
				# add to row count
				row_sum+=cell

		row_sums[r] = float(row_sum)

	return row_sums


def combine_sums(max_rows, all_sums):
	row_counts = [0] * max_rows

	# go through every list in all_row_sums
	for sums in all_sums:
		i = 0
		# go through every number in the list and add it to the matching index in row_counts
		for n in sums:
			row_counts[i] += sums[i]
			i+=1
	
	return row_counts


# Find the sum of all rows in an excel workbook.
# Returns a list of row sums, where each index represents a row number,
# and each list value represents a row sum.
# Note: this doesn't include the first row of headers. (Q: Should it?)
def sum_wb_rows(wb):

	# get list of all sheets in workbook
	wb_ = pd.ExcelFile(wb)
	all_sheets = wb_.sheet_names

	all_row_sums = [0]*len(all_sheets)
	max_rows = 0

	# for each sheet in the workbook, sum the rows
	for i in range(len(all_sheets)):
		df = pd.read_excel(wb, sheet_name=all_sheets[i])
		row_sums = sum_sheet_rows(df)
		all_row_sums[i] = row_sums # store list of row sums in an array

		if len(row_sums) > max_rows: # keep track of the max number of rows
			max_rows = len(row_sums) # for all sheets in workbook

	row_counts = combine_sums(max_rows, all_row_sums)
	
	return row_counts

# Find the sum of all rows of all sheets of all Excel workbooks in a list
# of Excel workbooks. 
# Takes in a list of Excel workbook file paths.
# Returns a list of row sums, where each index represents a row number,
# and each list value represents a row sum.
def sum_rows(wb_list):

	max_rows = 0
	all_sums = [0]*len(wb_list)

	for i in range(len(wb_list)):
		sums = sum_wb_rows(wb_list[i])
		all_sums[i] = sums

		if len(sums) > max_rows:
			max_rows = len(sums)

	total_sums = combine_sums(max_rows, all_sums)

	return total_sums


def sum_excel(file_directory):

	find_workbooks(file_directory)

	print(workbooks)

	row_count_list = sum_rows(workbooks)

	print(row_count_list)

	# make a csv file to store sums of excel rows
	sums_file = open('row_sums.csv', 'w')
	with sums_file:
		for num in row_count_list:
			writer = csv.writer(sums_file)
			writer.writerow([num])


# Testing
sum_excel(".")

