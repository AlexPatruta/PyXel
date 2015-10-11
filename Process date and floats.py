'''
Created on 11 oct. 2015

@author: Alex
'''
import xlrd, datetime
excel_file = "test.xlsx"

workbookObj = xlrd.open_workbook(excel_file)

sheetObj = workbookObj.sheet_by_name("TEST")

if __name__ == '__main__':
    #print sheetObj
    
    print("Rows: "), sheetObj.nrows
    print("Cols: "), sheetObj.ncols
        
    data = [[sheetObj.cell_value(row, col) for col in range(sheetObj.ncols)] for row in range(sheetObj.nrows)]
    print("Data from excel before: ")
    for item in data:
        for entry in item:
            print("    "),entry,
        print("")
        
    #Transform the date, age and gender cell into a more pleasand format for printing
    try:
        print("Processing \"Date\" field...")
        for col in range(sheetObj.ncols):
            for row in range(sheetObj.nrows):
                if sheetObj.cell_type(row, col) == 3:
                    excel_date = sheetObj.cell_value(row, col)
                    time_tuple = xlrd.xldate_as_tuple(excel_date, 0)
                    data[row][col] = datetime.datetime(*time_tuple)
                    
        print("Processing \"Age\" field...")
        for col in range(sheetObj.ncols):
            for row in range(sheetObj.nrows):
                if type(data[row][col]) is float:
                    data[row][col] = int(data[row][col])
                    
        print("Processing \"Gender\" field...")
        for col in range(sheetObj.ncols):
            for row in range(sheetObj.nrows):
                if data[row][col] == "M":
                    data[row][col] = "Male"
                elif data[row][col] == "F":
                    data[row][col] = "Female"
                else:
                    pass
        
    finally:
        print("Conversion successfully.\nPrinting...")
        for item in data:
            for entry in item:
                print("    "),entry,
            print("")
    
    
    
    
    
    
    
