ExcelApplication.WorkBook.SaveAs(filename, FileFormat, Password,
WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode,
ConflictResolution, AddToMru, TextCodePage, TextVisualLayout, Local)

1、Filename: Variant 类型 ，可选 。该字符串表示要保存的文件名 。可包含完整路径 。
如果不指定路径 ，Microsoft Excel 将文件保存到当前文件夹中 。

2、FileFormat: Variant 类型 ，可选 。保存文件时使用的文件格式 。要得到有效选项的
列表 ，请参阅 FileFormat 属性 。对于已有文件 ，其默认格式是上次指定的文件格式 ；
对于新文件 ，默认格式为当前使用的 Excel 版本格式, 可选常数如下 ：

xlAddIn 18

xlCSV 6

xlCSVMac 22

xlCSVMSDOS 24

xlCSVWindows 23

xlCurrentPlatformText - 4158

xlDBF2 7

xlDBF3 8

xlDBF4 11

xlDIF 9

xlExcel2 16

xlExcel2FarEast 27

xlExcel3 29

xlExcel4 33

xlExcel4Workbook 35

xlExcel5 39

xlExcel7 39

xlExcel9795 43

xlHtml 44

xlIntlAddIn 26

xlIntlMacro 25

xlSYLK 2

xlTemplate 17

xlTextMac 19

xlTextMSDOS 21

xlTextPrinter 36

xlTextWindows 20

xlUnicodeText 42

xlWebArchive 45

xlWJ2WD1 14

xlWJ3 40

xlWJ3FJ3 41

xlWK1 5

xlWK1ALL 31

xlWK1FMT 30

xlWK3 15

xlWK3FM3 32

xlWK4 38

xlWKS 4

xlWorkbookNormal - 4143

xlWorks2FarEast 28

xlWQ1 34

xlXMLData 47

xlXMLSpreadsheet 46

3、Password: Variant 类型 ，可选 。它是一个区分大小写的字符串
（最长不超过 15 个字符 ），用于指定文件的保护密码 。

4、WriteResPassword: Variant 类型 ，可选 。该字符串表示文件的写保护密码 。如果
文件保存时带有密码 ，但打开文件时不输入密码 ，则该文件以只读方式打开 。

5、ReadOnlyRecommended: Variant 类型 ，可选 。如果该值为 True ，则在打开文件
时显示一条信息 ，提示该文件以只读方式打开 。

6、CreateBackup: Variant 类型 ，可选 。如果该值为 True ，则创建备份文件 。

7、AccessMode: XlSaveAsAccessMode 类型 ，可选, 如果省略此参数 ，则不会更改访问
方式 。如果在没有改变文件名的情况下保存共享列表 ，则可以忽略参数 。若要改变访问
方式 ，请使用 ExclusiveAccess 方法 。以下是可选常量 ：

xlExclusive 3 不含方式

xlNoChange 1 不更改访问方式, 缺省值

xlShared 2 共享列表

8、ConflictResolution: XlSaveConflictResolution 类型 ，如果忽略本参数 ，则显示冲
突处理对话框 。可选常量如下:

xlLocalSessionChanges 2 自动接受本地用户的修改

xlOtherSessionChanges 3 接受除本地用户之外的其他用户的更改

xlUserResolution 1 显示冲突解决方案对话框

9、AddToMru ：Variant 类型 ，可选 。如果该值为 True ，则将该工作簿添加到最近
使用的文件列表中 。默认值为 False 。

10、TextCodePage ：Variant 类型 ，可选 。不在美国英语版的 Microsoft Excel 中使用 。

11、TextVisualLayout ：Variant 类型 ，可选 。不在美国英语版的 Microsoft Excel 中
使用 。

12、Local ：Variant 类型 ，可选 。如果该值为 True ，则以 Microsoft Excel
（包括控制面板设置 ）的语言保存文件 。如果该值为 False （默认值 ），则以
Visual Basic For Applications(VBA) 的语言保存文件 ，其中 Visual Basic For
Applications(VBA) 为典型安装的美国英语版本 ，除非 VBA 项目中的
Workbooks.Open 来自旧的国际化的 XL5 / 95 VBA 项目





18 XlFileFormat.xlAddIn Microsoft Office Excel 加载宏( * .xla)
6 XlFileFormat.xlCSV CSV(逗号分隔)( * .csv)
22 XlFileFormat.xlCSVMac
24 XlFileFormat.xlCSVMSDOS
23 XlFileFormat.xlCSVWindows
 - 4158 XlFileFormat.xlCurrentPlatformText
7 XlFileFormat.xlDBF2 DBF 2(dBASE II)( * .dbf)
8 XlFileFormat.xlDBF3 DBF 3(dBASE III)( * .dbf)
11 XlFileFormat.xlDBF4 DBF 4(dBASE IV)( * .dbf)
9 XlFileFormat.xlDIF DIF(数据交换格式)( * .dif)
16 XlFileFormat.xlExcel2 Microsoft Excel 2.1 工作表( * .xls)
27 XlFileFormat.xlExcel2FarEast
29 XlFileFormat.xlExcel3 Microsoft Excel 3.0 工作表( * .xls)
33 XlFileFormat.xlExcel4 Microsoft Excel 4.0 工作表( * .xls)
35 XlFileFormat.xlExcel4Workbook Microsoft Excel 4.0 工作簿( * .xlw)
39 XlFileFormat.xlExcel5 Micorosoft Excel 5.0 / 95 工作薄( * .xls)
39 XlFileFormat.xlExcel7 Micorosoft Excel 5.0 / 95 工作薄( * .xls)
43 XlFileFormat.xlExcel9795 Microsoft Excel 97 - Excel 2003 & 5.0 / 95 工作簿(. * .xls)
44 XlFileFormat.xlHtml 网页( * .htm ; * .html)
26 XlFileFormat.xlIntlAddIn
25 XlFileFormat.xlIntlMacro
2 XlFileFormat.xlSYLK SYLK(符号链接)( * .slk)
17 XlFileFormat.xlTemplate 模板( * .xlt)
19 XlFileFormat.xlTextMac
21 XlFileFormat.xlTextMSDOS 文本文件(制表符分隔)( * .txt)
36 XlFileFormat.xlTextPrinter 带格式文本文件(空格分隔)( * .prn)
20 XlFileFormat.xlTextWindows
42 XlFileFormat.xlUnicodeText Unicode 文本( * .txt)
45 XlFileFormat.xlWebArchive 单个文件网页( * .mht ; * .mhtml)
14 XlFileFormat.xlWJ2WD1 WD1(1 - 2 - 3)( * .wd1)
40 XlFileFormat.xlWJ3
41 XlFileFormat.xlWJ3FJ3
5 XlFileFormat.xlWK1 WK1(1 - 2 - 3)( * .wk1)
31 XlFileFormat.xlWK1ALL WK1, ALL(1 - 2 - 3)( * .wk1)
30 XlFileFormat.xlWK1FMT WK1, FMT(1 - 2 - 3)( * .wk1)
15 XlFileFormat.xlWK3 WK3(1 - 2 - 3)( * .wk3)
32 XlFileFormat.xlWK3FM3 WK3, FM3(1 - 2 - 3)( * .wk3)
38 XlFileFormat.xlWK4 WK4(1 - 2 - 3)( * .wk4)
4 XlFileFormat.xlWKS WKS(Works)( * .wks)
 - 4143 XlFileFormat.xlWorkbookNormal Microsoft Office Excel 工作簿( * .xls)
28 XlFileFormat.xlWorks2FarEast
34 XlFileFormat.xlWQ1 WQ1(Quattro Pro / DOS)( * .wq1)
46 XlFileFormat.xlXMLSpreadsheet XML 表格( * .xml)