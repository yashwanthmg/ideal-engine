# import xlsxwriter module
import xlsxwriter
import datetime
import tkinter as tk
from tkinter import simpledialog

from datetime import datetime


# Workbook() takes one, non-optional, argument
# which is the filename that we want to create.
workbook = xlsxwriter.Workbook('QL file.xlsx')

# The workbook object is then used to add new
# worksheet via the add_worksheet() method.
worksheet = workbook.add_worksheet()

f1 = workbook.add_format()
f1.set_bold(True)

border_format=workbook.add_format({'border':1})
worksheet.conditional_format( 'A1:E1' , { 'type' : 'no_blanks' , 'format' : border_format} )

date_format = workbook.add_format({'num_format':'dd-mm-yyyy'})


ROOT = tk.Tk()

ROOT.withdraw()

USER_INP1 = simpledialog.askstring(title="Create QL File", prompt="Enter ASIN:")
USER_INP2 = simpledialog.askstring(title="Create QL File", prompt="Seller ID:")
USER_INP3 = datetime.datetime.now() + datetime.timedelta(days=-1)
USER_INP4 = datetime.datetime.now() + datetime.timedelta(days=1)
USER_INP5 = simpledialog.askstring(title="Create QL File", prompt="Quantity Limit:")
USER_INP6 = simpledialog.askstring(title="Create QL File", prompt="SIM No:")
USER_INP7 = simpledialog.askstring(title="Create QL File", prompt="Business Name:")
USER_INP8 = simpledialog.askstring(title="Create QL File", prompt="Email ID:")



# Use the worksheet object to write
# data via the write() method.
worksheet.write('A1', 'ASIN',f1)
worksheet.write('B1', 'Seller ID',f1)
worksheet.write('C1', 'Start Date',f1)
worksheet.write('D1', 'End Date',f1)
worksheet.write('E1', 'Quantity Limit',f1)



# check it out
worksheet.write('A2', USER_INP1)
worksheet.write('B2', USER_INP2)
worksheet.write('C2', USER_INP3)
worksheet.write('D2', USER_INP4)
worksheet.write('E2', USER_INP5)


# Finally, close the Excel file
# via the close() method.
workbook.close()


workbook = xlsxwriter.Workbook('Revoke File.xlsx')

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
worksheet.write('B1', 'Seller ID',f1)
worksheet.write('C1', 'Start Date',f1)
worksheet.write('D1', 'End Date',f1)
worksheet.write('E1', 'Quantity Limit',f1)


ROOT = tk.Tk()

ROOT.withdraw()
# the input dialog
USER_INP3 = Start Date
USER_INP4 = End Date
USER_INP5 = simpledialog.askstring(title="Create QL File", prompt="Quantity Limit:")
# check it out
worksheet.write('A2', USER_INP)
worksheet.write('B2', USER_INP2)
worksheet.write('C2', USER_INP3)
worksheet.write('D2', USER_INP4)
worksheet.write('E2', USER_INP5)



# Finally, close the Excel file
# via the close() method.
workbook.close()


workbook = xlsxwriter.Workbook('Revoke File.xlsx')

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

from datetime import datetime

date_format = workbook.add_format({'num_format':'yyyy-mm-dd hh:mm:ss'})

worksheet.write(0, 0, datetime.today(),date_format)

ROOT = tk.Tk()

ROOT.withdraw()
# the input dialog
USER_INP1 = simpledialog.askstring(title="Create QL File", prompt="Enter ASIN:")
USER_INP2 = simpledialog.askstring(title="Create QL File", prompt="OfferType:")
USER_INP3 = simpledialog.askstring(title="Create QL File", prompt="StartDate:")
USER_INP4 = simpledialog.askstring(title="Create QL File", prompt="EndDate:")

# check it out
worksheet.write('A2', USER_INP)
worksheet.write('B2', USER_INP2)
worksheet.write('C2', USER_INP3)
worksheet.write('D2', USER_INP4)



# Finally, close the Excel file
# via the close() method.
workbook.close()
