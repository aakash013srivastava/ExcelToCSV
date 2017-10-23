from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
import xlrd
import csv

input_filename = None
output_directory = None

root = Tk()

lbl_name = Label(root,text = "Excel File Path" )
lbl_name.grid(row = 0,sticky = E)
lbl_pass = Label(root,text = "Output CSV Path")
lbl_pass.grid(row = 1)

def openinputfile():
	global input_filename
	input_filename = askopenfilename(filetypes=(
											   ("Excel files", "*.xls;*.xlsx"),
											   ("All files", "*.*") ))
	print (input_filename)

def openoutputdirectory():
	global output_directory
	output_directory = askdirectory()
	print (output_directory)
	
input_browse = Button(text = "Browse",command = openinputfile)
input_browse.grid(row= 0,column = 1)

output_browse = Button(text = "Browse",command = openoutputdirectory)
output_browse.grid(row= 1,column = 1)

										   
def convertxls():
	
	wb = xlrd.open_workbook(input_filename)
	sh = wb.sheet_by_name('Sheet1')
	your_csv_file = open(output_directory+'\\your_csv_file.csv', 'w')
	wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
	for rownum in range(sh.nrows):
		wr.writerow(sh.row_values(rownum))
	your_csv_file.close()

button1 = Button(text = "Convert",bg = "red",fg = "black",command=convertxls)
button1.grid(row = 4,columnspan = 2)



root.mainloop()