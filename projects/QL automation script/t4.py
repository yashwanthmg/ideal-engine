# import xlsxwriter module
import xlsxwriter
import datetime
import tkinter as tk
import pyperclip
from pytz import timezone

asin = None
quantityrequired = None
startdate = None
enddate = None
sellerid = None
sellername = None
offertype = None
quantitylimit = None
simid = None
businessname = None
emailid = None

file_name1= None
file_name2= None
file_name3= None


root = tk.Tk()
root.geometry("520x500")

def update_label():
	asin = asin_var.get()
	quantityrequired = quantityrequired_var.get()
	sellerid = sellerid_var.get()
	sellername = sellername_var.get()
	businessname = businessname_var.get()
	emailid = emailid_var.get()
	
	label11.config(text=f"Business Name: {businessname} \n\nEmail ID: {emailid} \n\nASIN: {asin} \n\nQTY Required: {quantityrequired} \n\nSeller ID: {sellerid} \n\nSeller Name: {sellername}")

def get_input():
	global asin 
	asin = asin_var.get()
	global quantityrequired
	quantityrequired = quantityrequired_var.get()
	global sellerid 
	sellerid = sellerid_var.get()
	global sellername
	sellername = sellername_var.get()
	global quantitylimit 
	quantitylimit = quantitylimit_var.get()
	global businessname 
	businessname = businessname_var.get()
	global emailid 
	emailid = emailid_var.get()
	global offertype 
	offertype = offertype_var.get()
	global associatename 
	associatename = associatename_var.get()
	global simid 
	simid = simid_var.get()
	
	global file_name
	DATE_INP4 = str(datetime.datetime.now().date())
	file_name1 = simid + '+' + associatename + '+' + sellername + '+' + offertype + '+' + DATE_INP4 + '+' +'QL template' + '.xlsx'
	file_name2 = simid + '+' + associatename + '+' + sellername + '+' + offertype + '+' + DATE_INP4 + '+' + 'Revoke template' + '.xlsx'
	file_name3 = simid + '+' + associatename + '+' + sellername + '+' + offertype + '+' + DATE_INP4 + '+' +'Time template' + '.xlsx'
	global root
	root.quit()
	
	# Workbook() takes one, non-optional, argument
	# which is the filename that we want to create.
	workbook = xlsxwriter.Workbook(file_name1)

	# The workbook object is then used to add new
	# worksheet via the add_worksheet() method.
	worksheet = workbook.add_worksheet()

	f1 = workbook.add_format()
	f1.set_bold(True)

	border_format=workbook.add_format({'border':1})
	worksheet.conditional_format( 'A1:E1' , { 'type' : 'no_blanks' , 'format' : border_format} )
	date_format = workbook.add_format({'num_format':'dd-mm-yyyy'})
	
	DATE_INP1 = datetime.datetime.now().date() + datetime.timedelta(days=-1)
	DATE_INP2 = datetime.datetime.now().date() + datetime.timedelta(days=1)
	
	# Use the worksheet object to write
	# data via the write() method.
	worksheet.write('A1', 'ASIN',f1)
	worksheet.write('B1', 'Seller ID',f1)
	worksheet.write('C1', 'Start Date',f1)
	worksheet.write('D1', 'End Date',f1)
	worksheet.write('E1', 'Quantity Limit',f1)


	# check it out
	worksheet.write('A2', asin)
	worksheet.write('B2', sellerid)
	worksheet.write('C2', DATE_INP1,date_format)
	worksheet.write('D2', DATE_INP2,date_format)
	worksheet.write('E2', quantityrequired)


	# Finally, close the Excel file
	# via the close() method.
	workbook.close()

	# Workbook() takes one, non-optional, argument
	# which is the filename that we want to create.
	workbook = xlsxwriter.Workbook(file_name2)

	# The workbook object is then used to add new
	# worksheet via the add_worksheet() method.
	worksheet = workbook.add_worksheet()

	f1 = workbook.add_format()
	f1.set_bold(True)

	border_format=workbook.add_format({'border':1})
	worksheet.conditional_format( 'A1:E1' , { 'type' : 'no_blanks' , 'format' : border_format} )
	date_format = workbook.add_format({'num_format':'dd-mm-yyyy'})
	
	DATE_INP3 = datetime.datetime.now().date()
	DATE_INP4 = '31-12-2023'
	
	# Use the worksheet object to write
	# data via the write() method.
	worksheet.write('A1', 'ASIN',f1)
	worksheet.write('B1', 'Seller ID',f1)
	worksheet.write('C1', 'Start Date',f1)
	worksheet.write('D1', 'End Date',f1)
	worksheet.write('E1', 'Quantity Limit',f1)


	# check it out
	worksheet.write('A2', asin)
	worksheet.write('B2', sellerid)
	worksheet.write('C2', DATE_INP3,date_format)
	worksheet.write('D2', DATE_INP4,date_format)
	worksheet.write('E2', quantitylimit)


	# Finally, close the Excel file
	# via the close() method.
	workbook.close()
	
	# Workbook() takes one, non-optional, argument
	# which is the filename that we want to create.
	workbook = xlsxwriter.Workbook(file_name3)

	# The workbook object is then used to add new
	# worksheet via the add_worksheet() method.
	worksheet = workbook.add_worksheet()

	f1 = workbook.add_format()
	f1.set_bold(True)

	border_format=workbook.add_format({'border':1})
	worksheet.conditional_format( 'A1:E1' , { 'type' : 'no_blanks' , 'format' : border_format} )

	
	# Use the worksheet object to write
	# data via the write() method.
	worksheet.write('A1', 'ASIN',f1)
	worksheet.write('B1', 'OfferType',f1)
	worksheet.write('C1', 'StartDate',f1)
	worksheet.write('D1', 'EndDate',f1)


	# check it out
	worksheet.write('A2', asin)
	worksheet.write('B2', offertype)



	# Finally, close the Excel file
	# via the close() method.
	workbook.close()
	
def copy_text():
    pyperclip.copy(label11["text"])
	
asin_var = tk.StringVar()
asin_var.trace("w", lambda *args: update_label())

label1 = tk.Label(root, text="ASIN")
label1.grid(row=0, column=0, sticky="W", padx=10, pady=10)
entry1 = tk.Entry(root, textvariable=asin_var)
entry1.grid(row=0, column=1, padx=10, pady=10)

quantityrequired_var = tk.StringVar()
quantityrequired_var.trace("w", lambda *args: update_label())

label2 = tk.Label(root, text="QTY Required")
label2.grid(row=1, column=0, sticky="W", padx=10, pady=10)
entry2 = tk.Entry(root, textvariable=quantityrequired_var)
entry2.grid(row=1, column=1, padx=10, pady=10)

sellerid_var = tk.StringVar()
sellerid_var.trace("w", lambda *args: update_label())

label3 = tk.Label(root, text="Seller ID")
label3.grid(row=0, column=2, sticky="W", padx=10, pady=10)
entry3 = tk.Entry(root, textvariable=sellerid_var)
entry3.grid(row=0, column=3, padx=10, pady=10)

sellername_var = tk.StringVar()
sellername_var.trace("w", lambda *args: update_label())

label4 = tk.Label(root, text="Seller Name")
label4.grid(row=1, column=2, sticky="W", padx=10, pady=10)
entry4 = tk.Entry(root, textvariable=sellername_var)
entry4.grid(row=1, column=3, padx=10, pady=10)

businessname_var = tk.StringVar()
businessname_var.trace("w", lambda *args: update_label())

label5 = tk.Label(root, text="Business Name")
label5.grid(row=2, column=0, sticky="W", padx=10, pady=10)
entry5 = tk.Entry(root, textvariable=businessname_var)
entry5.grid(row=2, column=1, padx=10, pady=10)

emailid_var = tk.StringVar()
emailid_var.trace("w", lambda *args: update_label())

label6 = tk.Label(root, text="Email ID")
label6.grid(row=3, column=0, sticky="W", padx=10, pady=10)
entry6 = tk.Entry(root, textvariable=emailid_var)
entry6.grid(row=3, column=1, padx=10, pady=10)

associatename_var = tk.StringVar()
associatename_var.trace("w", lambda *args: update_label())

label7 = tk.Label(root, text="Associate Name")
label7.grid(row=4, column=0, sticky="W", padx=10, pady=10)
entry7 = tk.Entry(root, textvariable=associatename_var)
entry7.grid(row=4, column=1, padx=10, pady=10)

quantitylimit_var = tk.StringVar()
quantitylimit_var.trace("w", lambda *args: update_label())

label8 = tk.Label(root, text="Quantity Limit")
label8.grid(row=2, column=2, sticky="W", padx=10, pady=10)
entry8 = tk.Entry(root, textvariable=quantitylimit_var)
entry8.grid(row=2, column=3, padx=10, pady=10)

offertype_var = tk.StringVar()
offertype_var.trace("w", lambda *args: update_label())

label9 = tk.Label(root, text="Offer Type")
label9.grid(row=3, column=2, sticky="W", padx=10, pady=10)
entry9 = tk.Entry(root, textvariable=offertype_var)
entry9.grid(row=3, column=3, padx=10, pady=10)

simid_var = tk.StringVar()
simid_var.trace("w", lambda *args: update_label())

label10 = tk.Label(root, text="Sim ID")
label10.grid(row=4, column=2, sticky="W", padx=10, pady=10)
entry10 = tk.Entry(root, textvariable=simid_var)
entry10.grid(row=4, column=3, padx=10, pady=10)

label11 = tk.Label(root, text="")
label11.grid(row=6, column=1, columnspan=2, pady=10)

copy_button = tk.Button(root, text="Copy", command=copy_text)
copy_button.grid(row=7, column=1, padx=10, pady=10)

button = tk.Button(root, text="Submit", command=get_input)
button.grid(row=7, column=2, pady=10)


root.mainloop()
