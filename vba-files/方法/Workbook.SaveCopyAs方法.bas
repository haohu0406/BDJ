Workbook.SaveCopyAs 方法(Excel)

语法
参数
示例
将工作簿副本保存到文件中 ，但不修改内存中打开的工作簿 。
语法
表达式 。SaveCopyAs(FileName)
_表达式_一个代表 * *工作簿 * *对象的变量 。
参数
参数
名称 必需 / 可选 数据类型 说明
FileName 必需 Variant 指定副本的文件名称 。
示例
本示例保存活动工作簿的副本 。
VB

复制
ActiveWorkbook.SaveCopyAs "C:\TEMP\XXXX.XLS"