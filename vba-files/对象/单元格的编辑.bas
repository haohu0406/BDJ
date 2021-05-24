1.单元格的输入
range("") = "" & ""
range("b2") = "a" & char(10) & "b" 'char(10)换行符

2.单元格的复制剪切
range("a1:10").copy range("c1") 'a1:a10的内容复制到C1

range("a1:a10").copy
activesheet.paste range("d1")

range("a1:a10").copy
range("e1").pastespecial(xlpastevalues) '选择性粘贴为数值

range("a1:a10").cut

range("a1:a10").copy
range("c1:c10").pastespecial opration:=xladd '选择性粘贴-加

range("a1:a10") = range("c1:c10").value

3.填充公式
range("a1") = "a1*10"
range("a1:a10").filldown

4.插入并复制公式
rows(4).insert
range("3:4").filldown
range("4:4").specialcells(xlcelltypeconstants) = “”'定位常量单元格赋值

columns(1).specialcells(xlcelltypeblanks).entirerow.delete '删除空值所在的行

5.删除单元格的内容 ，保留格式
range("a1").ClearContents