1 字典的遍历方法
即使是后期绑定 ，也可以用
For Each i In dict1.keys() '这里用到的 dict1.keys()   但不是 dict1.keys(i) 是可以的
    Sub test_dict11()
        
        Dim dict1 As Object
        Set dict1 = CreateObject("scripting.dictionary")
        
        
        dict1.Add 1, "h"
        dict1.Add 2, "e"
        dict1.Add 3, "l"
        dict1.Add 4, "l"
        dict1.Add 5, "o"
        dict1.Add 6, "world"
        
        For Each I In dict1.keys()
            Debug.Print I & "," & dict1(I)
        Next
        Debug.Print
        '知道key 去查item
        
        Debug.Print dict1(1)
        Debug.Print dict1.Item(2)
        Debug.Print
        
        
        
    End Sub
    
    
    2 用key查item ，常规操作
    dict1(key)
    dict1.items(key)
    
    
    3 用item查key 非常规操作
    3.1 我自己的思路 ：dict1.keys() 和 dict1.items() 是一对对对应的
    利用key, item是成对出现的 字典设计
    如果item不重复的话 ，那么item的index就是key的index ，index是同步的
    即使item重复 ，也可以循环筛出多个index
    但是要注意 ，这个方法必须前期绑定才生效 。因为只有前期绑定 ，才支持 dict1.keys(i) 这种循环
    
    
    不减1就出错 ，因为match函数的问题 ，是工作表函数
    这样用match查找 ，还是需要前期绑定的支持 ，因为 dict1.keys(index) 直接用必须前期绑定
    Sub test_dict1()
        
        Dim dict1 As New Dictionary
        
        dict1.Add 1, "h"
        dict1.Add 2, "e"
        dict1.Add 3, "l"
        dict1.Add 4, "l"
        dict1.Add 5, "o"
        dict1.Add 6, "world"
        
        '方法1
        '下面的反查方法要生效，必须是前期绑定！！！
        '知道item去查key
        '利用key,item 应该是成对的，也就是分别在 keys() items()的index是一样的。
        For J = LBound(dict1.Items()) To UBound(dict1.Items())
            If dict1.Items(J) = "world" Then
                '     If dict1(J) = "h" Then
                Debug.Print J 'keys(),items()数组没定义,index是从0开始的
                Debug.Print dict1.Keys(J)
            End If
            
        Next
        Debug.Print
        
        '方法2，也只有前期绑定支持
        '不减1就出错，因为match函数的问题，是工作表函数
        '这样用match查找，还是需要前期绑定的支持，因为 dict1.keys(index) 直接用必须前期绑定
        Debug.Print dict1.Keys(Application.Match("world", dict1.Items(), 0) - 1)
        
        
    End Sub
    
    
    
    
    
    
    4 方法2 ，后期绑定也生效
    后期绑定也可以查
    如果是后期绑定
    不能直接用 dict.keys() dict.items()
    但可以用个数组中转绕一下即可
    arr1 = dict.keys()
    arr1(index)
    arr2 = dict.items()
    arr2(index)
    
    
    Sub test_dict22()
        '后期绑定各种测试
        
        
        Dim dict1 As Object
        Set dict1 = CreateObject("scripting.dictionary")
        
        
        dict1.Add 1, "h"
        dict1.Add 2, "e"
        dict1.Add 3, "l"
        dict1.Add 4, "l"
        dict1.Add 5, "o"
        dict1.Add 6, "world"
        
        
        '后期绑定，为什么又可以使用dict1.keys()? for each可以使用？ 不能用index方式使用这2个数组？
        For Each I In dict1.Keys()
            Debug.Print I & " " ;
            Debug.Print dict1(I)
        Next
        Debug.Print
        
        
        
        '数组中转
        'dict1.keys()  dict1.items() 都可以使用，但不能做index for i 的这种遍历，中转的数组可以
        'dict1里 key-item同步，只要item不重复，可以用相同的变量循环
        
        arr1 = dict1.Keys()
        arr2 = dict1.Items()
        For I = LBound(arr2) To UBound(arr2)
            If arr2(I) = "world" Then
                Debug.Print arr1(I) & "," & arr2(I)
            End If
        Next
        Debug.Print
    End Sub
    
    
