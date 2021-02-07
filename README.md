使Excel分级目录进行快速自动分组，以方便分类查看。

使用xlwings的目的，xlwings可以很好的使用Office Com组件。使用xlwings的Obj对象的api就可以直接使用Com组件，非常方便。

Office官网中的Range("A1:G37").ClearOutline，需要换算为Range("A1:G37").ClearOutline()进行调用。
https://docs.microsoft.com/zh-cn/office/vba/api/overview/excel/object-model
https://docs.xlwings.org/en/stable/api.html#range

使用之前，请按实际环境，修改此三个变量。
path=r"D:\example.xlsx"
sheet_name  = 'Sheet1'
group_range = 'C2:G61'
