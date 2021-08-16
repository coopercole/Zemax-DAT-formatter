TITLE
Zemax .DAT formatter

DESCRIPTION
2dat.py is a file which can be configured to transfer data from an xlsx file, from 2 worksheets (each polarization), 
in that xlsx, and putting them into 1 worksheet within a workbook, formatted for a .dat file in Zemax.

REQUIRED
- to run use this formatter, the 2dat.py file and a .xlsx file to read the data from is required. 
- the code uses openpyxl, pandas, and string libraries

HOW TO USE
1. have the 2dat.py file, and the xlsx file you wish to read from in the same folder
2. install all required python libraries
3. open the 2dat.py file in an editor or ide
4. title the worksheet you will write to, using ws.title=""
5. title the workbook you will read from, using wbr = load_workbook(filename = '')
6. write the names of the of the 2 worksheets you wish to read from within the 'read' xlsx file 
7. title the table of the dat, using ws['B1'] = ""
8. name the workbook you wish to save to, using wb.save('')
9. ensure that the data within the read file, is in the same rows and columns that 2dat.py will be reading from
	9a. the reading happens under # Ts and # Tp, at (sheet_ranges1S["cell"].value) and (sheet_ranges1P["cell"].value)
	9b. there is a lot of tedious arithmetic that happens to ensure the cells all match up; using row2, col, and y variables
	9c. if you wish to change the read and write cells, be cautious. to be safe, leave 2dat.py as is, and change the read file to match read.xlsx (with data from C15 to U415)
10. with everything formatted, run 2dat.py, and the data will appear in the output.xlsx file, correctly formatted
11. to go from .xlsx to .dat, you will have to save the output.xlsx as a .txt file, then save that as a .dat file.
12. instead, you will probably copy the data from the .txt file to the master .DAT file 

NOTES
- the code in 2dat.py is commented to describe each line of code
- see example_output.txt, read.xlsx, output.xlsx for examples
- there is also the option to write all data from many different filters within one workbook, to a single workbook
-- see "additional.py" for sytanx or see below
-- format the names accordingly and use 	
	wb = load_workbook(filename='output.xlsx')
	ws = wb.create_sheet("sheetname")
	wbr = load_workbook(filename = 'read.xlsx')
	and at the end wb.save('output.xlsx')