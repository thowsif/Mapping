import openpyxl 
from openpyxl.styles import PatternFill


def main():
    book = openpyxl.load_workbook('newSheet2.xlsx')
    sheet = book.get_sheet_by_name('Sheet1')
    column_pos = 'G1'
    warnlist = openpyxl.load_workbook('warninglist.xlsx')
    warnlistSheet = warnlist.get_sheet_by_name('Sheet1')
    # print("rows:",warnlistSheet.max_row)

    # list_val = [3,4,14,24,35,41]
    # list_dict = {3:"COVERED" , 4:"COVERED" , 14:"NA",24:"Atlas",35:"Covered",41:"NA as per specs"}
    # list_val = list_val.sort()

    # print(list_val)
    
    sheet[column_pos].value = "REQ_ID"
    
    print(sheet['G1'].value)

    print("Rows:",sheet.max_row)
    print("Columns:",sheet.max_column)

    i=2
    id = 1
    string = "REQ_GML_WRN_000"
    for rows in range(1,sheet.max_row+1):
        for columns in range(1,sheet.max_column):
            if i< warnlistSheet.max_row+1 and warnlistSheet.cell(row=i,column=1).value == sheet.cell(row=rows,column=1).value:
                sheet.cell(row=rows,column=columns).fill = PatternFill(start_color="40e0d0",end_color="40e0d0",fill_type="solid")
                print("rows:",type(rows),columns)
                # print("max column:",sheet.max_column)
                if columns == sheet.max_column -1 :
                    # print("Columns:",columns)
                    # print("dict:",list_dict.get(list_val[i]))

                    # if list_dict.get(list_val[i]).upper() == "COVERED":
                    if warnlistSheet.cell(row=i,column=2).value=="" or warnlistSheet.cell(row=i,column=2).value.upper() == "COVERED":
                        sheet.cell(row=rows,column=columns+1).value = string + str(id)
                        id=id+1
                    else:
                        sheet.cell(row=rows,column=columns+1).value = warnlistSheet.cell(row=i,column=2).value
                    
                    # string = string + 1
                    i=i+1

            else:
                sheet.cell(row=rows,column=columns).fill = PatternFill(start_color="008000",end_color="008000",fill_type="solid")
            
            # print(sheet.cell(row=rows,column=columns).value)

    

    book.save('newSheet2.xlsx')


if __name__ == '__main__':
    main()    