定义全局变量可以去掉重复声明 ，但容易造成程序耦合
Public i As Integer
Dim i As Integer
放在页面顶端 ，其作用不同 ，public可以在其他模块引用 ，dim则只能在本模块中调用