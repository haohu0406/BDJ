我发现从word表格取出的内容最后是Chr(13) 即回车符CR和Chr(7) ，长度为2 ，
方框符号是Chr(7) 。用Replace(t, Chr(7), "") 替换后就可以正常显示了