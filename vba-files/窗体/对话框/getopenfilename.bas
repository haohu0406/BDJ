1、概述基本语法
getopenfilename相当于excel打开窗口 ，通过该窗口选择要打开的文件 ，并可以返回
选择的文件完整路径和文件名

application.getopenfilename(文件类型帅选规则 ，有限显示第几个类型的文件 ，
标题 ，是否允许选择多个文件名 ）

2.用法
'"文件类型说明文字，文件类型后缀"
Sub t1()
    Dim f
    f = application.getopenfilename("excel2003文件,*.xls,word文件,*.doc")
    MsgBox f
End Sub

3.打开多种文件类型 ，默认显示word文件
Sub t2
    Dim f
    f = application.geopenfilename("excel2003文件,*.xls,word文件,*.doc,文本文件,*.txt", 2)
    '2表示默认显示第二种文件
    MsgBox f
End Sub

4.设置对话框名称
Sub t4()
    Dim f
    f = application.getopenfilename("excel2003文件,*xls,word文件,*.doc,文本文件_
    ,  * .txt ",2," 选择要汇总的文件 ")
    MsgBox f
End Sub

5.选择多个文件 ，并以数组形式返回
Sub t5(0
    Dim f
    ChDrive "e" '默认打开盘符
    ChDir application.path '默认文件路径
    f = application.getopenfilename("excel2003文件,*.xls,word文件,*.doc,文本文件
     _ ,  * .txt ",1,multiselect:=true)
            MsgBox f(1)
End Sub