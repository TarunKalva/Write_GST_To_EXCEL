#!/usr/bin/env python
# coding: utf-8

# In[8]:


import xlrd
import xlwt
import os
from termcolor import colored
# TO MOVE TO THE REQUIRED DIRECTORY
os.chdir('C:/Users/abc/Desktop/SHEKAR 22-23') # Enter the path here and change '\' to '/'
path=os.getcwd()
# CREATE A NEW WORKBOOK AND THEN CREATE A NEW SHEET
new_wbk=xlwt.Workbook()
new_sheet=new_wbk.add_sheet('April')
list_of_files = os.listdir(path) # LISTDIR IS USED TO CONVERT THE FILES IN PATH TO LIST OF FILES
col=1   
style_string="font: bold on"
style = xlwt.easyxf(style_string)
new_sheet.write(0,0,'NO',style=style)
new_sheet.write(0,1,'DATE',style=style)
new_sheet.write(0,2,'COMPANY NAME',style=style)
new_sheet.write(0,3,'GSTIN',style=style)
new_sheet.write(0,4,'STATE',style=style)
new_sheet.write(0,5,'TAXABLE VALUE',style=style)
new_sheet.write(0,6,'IGST',style=style)
new_sheet.write(0,7,'CGST',style=style)
new_sheet.write(0,8,'SGST',style=style)
new_sheet.write(0,9,'TOTAL VALUE',style=style)
new_sheet.write(0,10,'QUANTITY',style=style)
new_sheet.write(0,11,'UNIT',style=style)
for i in list_of_files:
    workbook=xlrd.open_workbook(i)
    sheet=workbook.sheet_by_index(0)
    if int((sheet.cell_value(12,8).split(':')[1]).split('-')[1])!=4: #Enter the month number e.g: 4
        continue
    else:
        new_sheet.write(col,0,sheet.cell_value(10,0).split(':')[1])
        new_sheet.write(col,1,sheet.cell_value(11,0).split(':')[1].strip())
        new_sheet.write(col,2,sheet.cell_value(16,0).split(':')[1])
        if sheet.cell_value(21,0)=='SELF USE':
            new_sheet.write(col,3,sheet.cell_value(21,0).split(':'))
        else:
            new_sheet.write(col,3,sheet.cell_value(21,0).split(':')[1])
        new_sheet.write(col,4,sheet.cell_value(22,0).split(':')[1])
        new_sheet.write(col,5,sheet.cell_value(38,14))
        new_sheet.write(col,10,sheet.cell_value(26,5))
        new_sheet.write(col,11,sheet.cell_value(27,5))
        if sheet.cell_value(22,7)==36:
            new_sheet.write(col,7,sheet.cell_value(39,14))
            new_sheet.write(col,8,sheet.cell_value(40,14))
            new_sheet.write(col,9,sheet.cell_value(41,14))
        else:    
            new_sheet.write(col,6,sheet.cell_value(39,14))
            new_sheet.write(col,9,sheet.cell_value(40,14))
    col=col+1
new_wbk.save('gst1 details.xls') 


# In[2]:


print(list_of_files)


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




