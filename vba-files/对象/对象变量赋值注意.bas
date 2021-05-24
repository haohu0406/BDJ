给对象赋值 ，一定要加set

对象的方法赋值给对象则方法后一定要加 （）
例如 ：
Dim wb As workbook
Set wb = workbooks.open("d:/b.xls")
wb.close True '关闭的时候保存
wb.close False '关闭的额时候不保存