1.窗体名称
属性窗口改变

2.代码显示窗体
Sub 窗体显示()
    入库单.show
End Sub

3.窗体关闭
Private Sub userform_click() '窗体点击事件
    unload Me '点击后自动运行unload me语句用于关闭窗体
End Sub

4.窗体的显示
标题 ：属性 - caption
名称 ：和标题显示区别开来 ，名称可以在程序中引用
颜色 ：
窗体背景图片 ：属性 - picture ，位置大小 （picture下方的属性 ）
窗体大小 ：属性 - 位置 - （所有者中心 ，即excel文件中心 ）
"窗体显示中可编辑 ：属性 - showmodal （False ）"
"controltiptext属性  按钮提示文本"

5.控件
右键 - 附件控件 - 日期控件 （Date And Time ）
控件的对齐 ：选取所有 - 格式 - 统一尺寸 。。。
控件输入顺序的跳转 ：右键 - table键顺序
tablsestop禁止跳转
控件的循环 ：

'showmodel = False显示窗体的同时可以编辑单元格


Private Sub commandbutton1_click()
    Dim x As Integer
    For x = 0 To Me.controls.count - 1
        If TypeName(controls(x)) = "TextBox" Then '判断是否文本框
            'controls(x)第x个控件	'textbox区分字母大小写，请注意
        End If 'TypeName用函数判断一个控件的类型
    End Sub
    
    6. 窗体的导出导入
    
    
    标签
    wordwrap = True标签内容自动换行
    
    
    
    
    按钮控件
    'conrtoltiptext 鼠标放按钮上提示内容
    '按钮图片一般使用icon格式
    'visible属性 显示和隐藏按钮
    accelerator 按钮设置热键
    enable设置灰色 按钮不可用
    cancel = True 在按下esc时会运行
    Default  = true按回车键会运行
    
    tag 临时存放内容 ，控件小仓库
    Sub 显示tag值大于20的按钮()
        Dim x
        For x = 0 To contrlos.count - 1
            If Val(controls(x).tag) > 20 Then
                MsgBox controls(x).Name
            End If
        Next        
    End Sub
    
    1 个按钮执行多个程序
    Private Sub commandbutton1_click()
        If commandbutton1.caption = "打开" Then
            MsgBox "你已打开"
            commandbutton1.caption = "关闭"
        Else
            MsgBox "你已关闭"
            commandbutton1.caption = "打开"
        End If
    End Sub