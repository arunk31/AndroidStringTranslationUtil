
# coding: utf-8

# # String Res Generator

# ## Generating a string resource file for all available languages in the given excel file

# ### Used for parsing the translation file provide by a vendor

# #### Install the below dependency
# 
# Using Python 3
# 
# pip install xlrd

# In[25]:


import xlrd

# Open and read a workbook
book = xlrd.open_workbook(input()) # python <3 use raw_input() instead of input()
# Select the first worksheet
sh = book.sheet_by_index(0)

print("Sheet:{0}, Rows:{1} Cols:{2}".format(sh.name, sh.nrows, sh.ncols))

for cx in range(sh.ncols):
    if cx == 0 :
        #Contains only the keys
        continue;
    
    #Create the DOM structure and populate the data
    from xml.dom.minidom import getDOMImplementation
    impl = getDOMImplementation() 
    doc = impl.createDocument(None, "resources", None)
    root = doc.documentElement
    doc.appendChild(root)
    fileName = "d:/string_{0}.xml".format(cx)
    for rx in range(sh.nrows):
        if rx == 0:
            #Contains the column title
            fileName = "d:/string_{0}.xml".format(sh.row(rx)[cx].value)
            continue;
        # Create Element
        tempChild = doc.createElement("string")
        root.appendChild(tempChild)
        # Write Text
        nodeText = doc.createTextNode(sh.row(rx)[cx].value)
        tempChild.setAttribute( "name", sh.row(rx)[0].value)
        tempChild.appendChild(nodeText)
    #Write the document to a file
    doc.writexml( open(fileName, 'w'),
               indent="  ",
               addindent="  ",
               newl='\n')
    #Close the internal link to the file
    doc.unlink()

