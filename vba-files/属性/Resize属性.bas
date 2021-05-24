Range('A1').Resize(1, 3).Select'选中单元格A1：C1区域
Range('A1').Resize(2, 3).Select'选中A1到C2单元格区域
Range('A1').Resize(, 3).Select'选中单元格A1：C1区域
Range('A1').Resize(3).Select'选中单元格A1：A3区域
range("a1:c6").resize(6, 2) 返回的是range("a1:b6")
这种用法结合offset可以达到更好的效果 ：
sh.UsedRange.Offset(1, 0).Resize(, 12).Copy 复制工作表中使用的区域 ，下
移一行 ，默认行数 ，重选其中12列进行复制