Sub shellsort()
    Dim jg, i, j, temp, arr
    arr = [a1].CurrentRegion
    jg = 1
    If UBound(arr) > 13 Then
        Do While jg < UBound(arr)
            jg = jg * 3 + 1
        Loop
        jg = jg \ 9
    End If
    Do While jg
        For i = 1 + jg To UBound(arr)
            temp = arr(i, 1)
            For j = i - jg To 1 Step -jg
                If temp < arr(j, 1) Then
                    arr(j + jg, 1) = arr(j, 1)
                Else
                    Exit For
                End If
            Next
            arr(j + jg, 1) = temp
        Next
        jg = jg \ 3
    Loop
    Range("f1").Resize(UBound(arr)) = arr
End Sub
希尔排序

（1）希尔排序 （Shell sort ）这个排序方法又称为缩小增量排序 ，
是1959年D ·L ·Shell提出来的 。该方法的基本思想是 ：
设待排序元素序列有n个元素 ，首先取一个整数increment （小于n ）作为间隔
将全部元素分为increment个子序列 ，所有距离为increment的元素放在同一
个子序列中 ，在每一个子序列中分别实行直接插入排序 。然后缩小间隔
increment ，重复上述子序列划分和排序工作 。直到最后取increment = 1 ，
将所有元素放在同一个子序列中排序为止 。



（2）由于开始时 ，increment的取值较大 ，每个子序列中的元素较少 ，
排序速度较快 ，到排序后期increment取值逐渐变小 ，子序列中元素个数
逐渐增多 ，但由于前面工作的基础 ，大多数元素已经基本有序 ，所以排序
速度仍然很快 。