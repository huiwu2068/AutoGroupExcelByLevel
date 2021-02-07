# author : Hank Lee
# email: huiwu2068@gmail.com
import xlwings as xw
#使用xlwings的目的，是xlwings可以很好的使用Office Com组件。使用xlwings的Obj对象的api就可以直接使用Com组件，非常方便。
#Office官网中的Range("A1:G37").ClearOutline，需要换算为Range("A1:G37").ClearOutline()进行调用。
# https://docs.microsoft.com/zh-cn/office/vba/api/overview/excel/object-model
# https://docs.xlwings.org/en/stable/api.html#range

#使用之前，请修改此三个变量。
path=r"D:\example.xlsx"
sheet_name  = 'Sheet1'
group_range = 'C2:G61'

#判断此行是否可以分组，判断依据是同样行的之前的列的单元，是否有数据。
def determine_group(rng,row,start_column,end_column):
    for column in range(start_column,end_column):
        if rng(row,column).value!=None and rng(row,column).value != "":
            return False
    return True

def auto_group(rng):
    ws = rng.app
    start_column = 1
    for column in range(rng.columns.count,1,-1):
        #rng的范围已经从1开始了，不是sheet的Cell(1,1)，而是选定范围的1.
        #对于M~N,先分组第N列，再分组第N-1列，直到M列。
        start_row = 1
        end_row = 1
        for row in range(1,rng.rows.count):
            if determine_group(rng,row,start_column,column) == False:
                if start_row != end_row:
                    group_range = rng(start_row,column).get_address()+":"+rng(end_row-1,column).get_address()
                    #rng已经是单元格了，可以rng.api.Rows.Group(),但已经不可以rng(X,X).api.Rows.Group()，就算做也只是分组一个单元格
                    ws.range(group_range).api.Rows.Group()
                    start_row = end_row
                start_row += 1
                end_row += 1
            else:
                end_row += 1
        
        if start_row != end_row:
            group_range = rng(start_row,column).get_address()+":"+rng(end_row,column).get_address()
            ws.range(group_range).api.Rows.Group()
    return 

wb = xw.Book(path)
ws = wb.sheets[sheet_name]
ws.api.Rows.ClearOutline()
auto_group(ws.range(group_range))
wb.save()
