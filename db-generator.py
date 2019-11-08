#!/bin/env python
#
#  Copyright (c) 2019  European Spallation Source ERIC
#
#  The program is free software: you can redistribute
#  it and/or modify it under the terms of the GNU General Public License
#  as published by the Free Software Foundation, either version 2 of the
#  License, or any newer version.
#
#  This program is distributed in the hope that it will be useful, but WITHOUT
#  ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
#  FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for
#  more details.
#
#  You should have received a copy of the GNU General Public License along with
#  this program. If not, see https://www.gnu.org/licenses/gpl-2.0.txt
#
#
#   author  : Antoni Simelio
#   email   : Antoni.Simelio@esss.se
#   date    : August 01 02:57:42 CEST 2019
#   version : 0.3


try:
    from pip import main as pipmain
except ImportError:
    from pip._internal import main as pipmain

import xlrd
import sys



print ("Hello world");

# Checking the INPUT FILE
if ( len(sys.argv)<2 ):
	print ("ERROR : Enter the name of the Input EXCEL File")


print ("This is the name of the Input File: ", sys.argv[1])

MyFile = sys.argv[1]

#print ("Number of arguments: ", len(sys.argv))
#print ("The arguments are: " , str(sys.argv))


#file_location = "C:/Users/antonisimelio/Desktop/PYTHON/ISrc_LEBT.xlsx"
#workbook = xlrd.open_workbook(file_location)
workbook = xlrd.open_workbook(MyFile)
sheet = workbook.sheet_by_index(0)
#sheet.cell_value(0,0)



print ("Gathering of the data Started")

# Soul of the system
data = [[sheet.cell_value(r,c)
for c in range (sheet.ncols)]
for r in range (sheet.nrows)]


# Transformation of the fields in STRING
for c in range (sheet.ncols):
	for r in range (sheet.nrows):
		#Adata[c][r][0]=sheet.cell_value(r,c)
		data[r][c]=sheet.cell_value(r,c)
		
		if (data[r][c] == ''):
			data[r][c]=''
		else:
			if(type(data[r][c]) is int):
				#data[1][r][c]=1
				data[r][c]=str((data[r][c]))
			
			else:
				if(type(data[r][c]) is float):
					#Adata[1][r][c]=2
					data[r][c]=str(int(data[r][c]))
				else:
					if(type(data[r][c]) is str):
						data[r][c]=data[r][c]
					else:
						#data[r][c]='Special'
						data[r][c]=data[r][c]
				
		

print ("Gathering of the Data is Done")




	
print ("Opening the File") 
	
with open('Output_File.db','w') as f:
		
	
	# WRITING ON THE DB FILE
	print ("Writing on the File") 	
	
	# Checking all the ROWS
	for r in range (sheet.nrows):
	
		
		# Avoid the row OF TITLES --> The EXCEL FILE must START with PV
		if (data[r][0]!='PV'):
			
			#WRITE TITLE of the RECORD
			f.write('record("*", "'+data[r][0]+'")')
			f.write("\n{\n")			
		
			
			#WRITE FIELDS of the RECORD
			#Checking all the rows
			for c in range (sheet.ncols):
		
				# Avoid writing the Process Variables itself and the value Null
				if (data[0][c] !='PV' and data[r][c] !=''):
					f.write("    ")
					f.write('field('+data[0][c]+', "'+data[r][c]+'")')
					f.write("\n")
	
			# WRITE ENDING TITLE CLAUDATOR
			f.write("}\n\n")

print ("The writing process is finalized") 	
		


	
# THINGS TO FIX
# Create FUNCTIONS
# Name of the INPUT FILE [the 
# Name of the OUTPUT FILE -- by default it should work
# Change the little ' for "
# Adapt it to the PYTHON of the ESS	[the version is correct]
 
