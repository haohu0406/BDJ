expression.Open(FileName, UpdateLinks, ReadOnly, Format, Password,
WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable,
Notify, Converter, AddToMru, Local, CorruptLoad)

常用
workbooks.open(filename, ubdatelinks, readonly)
workbooks.open("a\d.xls", False, True) '只读方式打开，不更新链接