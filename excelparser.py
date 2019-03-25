print ("I can finally run")

import openpyxl 
# import Workbook
# wb = Workbook()

# grab the active worksheet
# ws = wb.active

# getting the first required excel file from the folder to compare folders
getFirstFile = openpyxl.load_workbook('C:\Ramisa\MIS\Test1.xlsx')

#the current sheet is active
sheetForFile1 = getFirstFile.active

#getting cell data from the file

for columnIndex in sheetForFile1:
    # getdata = 'A'+str(columnIndex)
    # print('getdata: '+getdata)
    
    getdata[columnIndex] = sheetForFile1['A1']
    # a = sheetForFile1['A2']
    # a3 = sheetForFile1.cell(row=3, column=1)

    print(getdata[columnIndex])
# print(a2.value) 
# print(a3.value)



# Data can be assigned directly to cells
# ws['A1'] = 42

# Rows can also be appended
# ws.append([1, 2, 3])

# Python types will automatically be converted
# import datetime
# ws['A2'] = datetime.datetime.now()

# Save the file
# wb.save("Test1.xlsx")

