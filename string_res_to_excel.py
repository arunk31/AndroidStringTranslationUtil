
# coding: utf-8

# # Excel Generator

# ## Generating an excel file from a string resource file

# ### Used for sharing with translation service provider

# #### Install the below dependency
# 
# Using Python 3
# 
# pip install XlsxWriter

# In[2]:


import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('d:/Translations.xlsx')
worksheet = workbook.add_worksheet()

#Get a file and parse it
from xml.dom import minidom
xmldoc = minidom.parse(input()) # python <3 use raw_input() instead of input()
itemlist = xmldoc.getElementsByTagName('string') 
print("Len : ", len(itemlist))

items = []

#Populate the data into a list
for s in itemlist :
    items.append([s.attributes['name'].value, s.firstChild.nodeValue])

# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

# Write the headers
worksheet.write(row, col,     "Keys")
worksheet.write(row, col + 1, "English")
row += 1

# Iterate over the data and write it out row by row.
for item, value in (items):
    worksheet.write(row, col,     item)
    worksheet.write(row, col + 1, value)
    row += 1

#Close the file
workbook.close()

