Application.CommandBars("Standard").Controls(2).BeginGroup = True
'在常用工具栏的第二个按钮前插入分隔符
'commandbar工具栏
'control命令


Application.CommandBars("命令按钮名称").Position = msoBarFloating
'使[命令按钮]悬浮在表格中
Application.CommandBars("命令按钮名称").Position = msoBarTop
'使[命令按钮]排列在工具栏中
Application.CommandBars("命令按钮名称").Position = msoBarTop
'使[命令按钮]排列在工具栏中