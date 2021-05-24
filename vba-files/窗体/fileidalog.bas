语法 ：applicaiton.filedialog(filedialogtype)

名称 值 描述
msofiledialogfilepicker 3 "文件选取器" 对话框
msofiledialogfolderpicker 4 "文件夹选取器" 对话框
msofiledialogopen 1 "打开" 对话框
msofiledialogsaveas 2 "另存为" 对话框


Sub test()
    Dim filepath As String
    On Error Resume Next    
    With application.filedialog(msofiledialogfilepicker)
         .show
        range("b3") =  .selecteditems(1)
        
    End With    
End Sub
'.SelectedItems(1)返回的是我们选取文件的路径
用法如下


Application.FileDialog(fileDialogType)
fileDialogType MsoFileDialogType 类型 ，必需 。文件对话框的类型 。
MsoFileDialogType 可为以下 MsoFileDialogType 常量之一 。

允许用户选择文件 。
msoFileDialogFilePicker

允许用户选择一个文件夹 。
msoFileDialogFolderPicker

允许用户打开文件 。
msoFileDialogOpen

允许用户保存一个文件 。
msoFileDialogSaveAs



分别举例如下 ：
1、msoFileDialogFilePicker

1）选择单个文件
代码:
Sub SelectFile()
    
    '选择单一文件
    With Application.FileDialog(msoFileDialogFilePicker)
         .AllowMultiSelect = False        
        '单选择        
         .Filters.Clear        
        '清除文件过滤器        
         .Filters.Add 'Excel Files', '*.xls;*.xlw'        
         .Filters.Add 'All Files', '*.*'       
        '设置两个文件过滤器
        If  .Show = -1 Then
            'FileDialog 对象的 Show 方法显示对话框，并且返回 -1（如果您按 OK）和 0（如果您按 Cancel）。
            MsgBox '您选择的文件是：' & .SelectedItems(1), vbOKOnly vbInformation, '提示'
        End If
    End With
End Sub

2）选择多个文件
代码:
Sub SelectFile()
    '选择多个文件
    Dim l As Long
    With Application.FileDialog(msoFileDialogFilePicker)
         .AllowMultiSelect = True
        '多选
         .Filters.Clear
        '清除文件过滤器
         .Filters.Add 'Excel Files', '*.xls;*.xlw'
         .Filters.Add 'All Files', '*.*'
        '设置两个文件过滤器
         .Show
        'FileDialog 对象的 Show 方法显示对话框，并且返回 -1（如果您按 OK）和 0（如果您按 Cancel）。
        For l = 1 To  .SelectedItems.Count
            MsgBox '您选择的文件是：' & .SelectedItems(l), vbOKOnly vbInformation, '提示'  
        Next
    End With
End Sub

2、msoFileDialogFolderPicker
代码:
Sub SelectFolder()
    '选择单一文件
    With Application.FileDialog(msoFileDialogFolderPicker)
        If  .Show = -1 Then
            'FileDialog 对象的 Show 方法显示对话框，并且返回 -1（如果您按 OK）和 0（如果您按 Cancel）。
            MsgBox '您选择的文件夹是：' & .SelectedItems(1), vbOKOnly vbInformation, '提示'
        End If
    End With
End Sub

文件夹仅能选择一个
3、msoFileDialogOpen
4、msoFileDialogSaveAs
使用方法与前两种相同
只是在.show
可以用.Execute方法来实际打开或者保存文件 。

FileDialog常用属性
Title 标题
Filter 设置过滤字符
FilterIndex 指定列表框中的默认的选项
AllowMultiSelect 是否多选