Find语法 ：

Range.Find(What, After, LookIn ，LookAt, SearchOrder, SearchDirection,
MatchCase, MatchByte, SearchFormat) （）

What ：唯一一个必选参数 ，含义为需要查询的内容 ，特点为变体类型 ，可以接受
数字 ，字符串 ，日期等数据类型

此外 ，该参数可以接受通配符 。

常用通配符

？一个任意字符 ，比如 ？a ？可以表示bac aaa等

！任意多个任意字符 （包括0个 ），比如 * a * 可以表示123aadd23 ，ade ，a等

MatchCase:是否匹配大小写字母 ，True - 大小写视作不同 ，False - 大小写视作相
同

LookAt ：匹配单元格 ，取值 ：1 代表单元格的内容必须与欲查找内容长度相同 ，不能
多出字符

2单元格的内容只需包含欲查找的字符串即可

SearchFormat ：是否按照格式查找 true是 ，false否

其他参数因为不常用就不一一介绍啦