range("e4").addcomment.Text "代头" & Chr(10) & "内容……" '添加批注 
range("e4").Comment.Visible = True '显示批注
ActiveSheet.UsedRange.Comment.Shape.TextFrame.AutoSize = True
'根据批注内容自动调整大小
ActiveSheet.UsedRange.ClearComments
'清除活动工作表已使用范围所有批注