只存在于thisworkbook和类模块中 ，功能 ：所有工作簿适用 ，
加载到宏中可以打开工作簿便适用
1.声明
Public withevents app As excel.application

Private Sub workbook.open() '把excel程序对象交给APP
    Set app = excel.application
End Sub
2个必须 ，必须先声明 ，其次必须把excel程序对象交给APP ，程序才能正常运行 。
其作用类似于set wb = thisworkbook, 然后才能将wb作为工作簿进行操作 。


2.程序实体
禁止使用宏的时候还是会提示能否打开
wb.close False '关闭不保存

Private Sub app_workbookbeforesave(ByVal wb As Object, ByVal saveasui As Boolean, cancel As Boolean)