#
# driving.py
#
import xlwings as xw
path=r"E:\01_Project\02_groupexcel\4.xlsx"

def determine_group(rng,row,start_column,end_column):
    for column in range(start_column,end_column):
        if rng(row,column).value!=None and rng(row,column).value != "":
            return False
    return True

def auto_group(ws,rng):
    start_column = rng.column
    for column in range(rng.last_cell.column,rng.column,-1):
        start_row = rng.row
        end_row = rng.row
        for row in range(rng.row,rng.last_cell.row):
            if determine_group(rng,row,start_column,column) == False:
                if start_row != end_row:
                    group_range = rng(start_row,column).get_address()+":"+rng(end_row-1,column).get_address()
                    ws.range(group_range).api.Rows.Group()
                    start_row = end_row
                start_row += 1
                end_row += 1
            else:
                end_row += 1
        
        if start_row != end_row:
            group_range = rng(start_row,column).get_address()+":"+rng(end_row-1,column).get_address()
            ws.range(group_range).api.Rows.Group()
    return 

wb = xw.Book(path)
ws = wb.sheets['Sheet1']
ws.api.Range("A1:E80").ClearOutline
auto_group(ws,ws.range('A2:E60'))
wb.save()
