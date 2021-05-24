1.选项按钮
每一组只能选一个
有框架的情况下分组则在每组中可选一个
没有框架则所有选项按钮只能选一个
选项按钮的值只有两个状态 ，True Or False
Private Sub optionbutton3_click()
    If Me.男.value = True Then
        
    End If
End Sub


2. 框架控件
将按钮分成很多组
3. 图像控件
4. 多页控件
设置多个控件页面


5.复选框
两个状态 True Or False
Private Sub commandbutton1_click()
    Dim sr As String
    If checkbox1.value = True Then sr = sr & "" & checkbox1.caption
    If checkbox2.value = True Then sr = sr & "" & checkbox2.caption
    If checkbox3.value = True Then sr = sr & "" & checkbox3.caption
    textbox3.value = sr
End Sub


6. 微调和滚动条
事件
滚动条事件
Private Sub scrollbar1_change()
    textbox1.value = scrollbar1.value    
End Sub
图片事件
Private Sub userform_initialize() '先加载图片文件夹中的图片名称
    Dim f As String
    f = Dir(thisworkbook.path & "\pic\*.jpg")
    Do
        Me.listbox1.additem f
        f = Dir
    Loop Until Len(f) = 0
    textbox2 = Date
End Sub
Private Sub listbox1_click()
    Dim path
    path = thisworkbook.path & "/pic/" & listbox1.value
    image1.picture = LoadPicture(path) '用loadpicture函数加载图片
End Sub

7.多页控件
页标题分行显示 multirow = True
value 从0开始的页面序号


