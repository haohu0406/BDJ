cells(2, 3).row 指单元格中在第几行 ，返回的是行号

Range(“A1:A10 ”).Rows.Count 运算的是区域A1: A10多少行

rows.count 最大行数


Rows(1) 和Rows("1:1") 有什么区别
rows(1) 指第一行 ，rows("1:1") 指行1 ，两者意思一样 ，但用法不一样 ，
1到10行可以用rows("1:10"), 但不可以用rows(1:10)