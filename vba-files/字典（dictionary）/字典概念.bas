key item
a 1
b 2
d 3
c 4
e 5

'有数组为何要学习字典
'原因：提速
'具体表现在：
'1）key列只能装入非重复的元素，利用这个特点可以很方便的提取不重复的值
'2）每一个对key对应一个唯一的item，只要指点key的值，就可以马上返回其对应的item
'利用字典可以实现快速查找

'字典有什么局限？
'1）字典只有两列，如果要处理多列的数据，还需要通过字符串的组合和拆分来实现
'2）字典调用会耗费一定的时间，如果是数据量不大，字典的有事就无法体现出来。
'3)字典不支持"yyyy/m/d"的日期格式，但支持"yyyy-mm-dd"的日期格式

'字典在哪里，如何创建？
'字典是由scrrun.dll链接库提供的，要调用字典有两种方法
'第一种方法：直接创建
Set d = CreatObject("scripting.dictionary")
'第二种方法：引用法
'工具-引用-浏览-找到scrrun.dll-确定