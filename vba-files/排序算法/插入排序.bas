Sub 插入排序()
    Dim i, j, arr, temp
    arr = [a1].currentregion
    For i = 2 To UBound(arr)
        temp = arr(i, 1)
        For j = i - 1 To 1 Step -1
            If temp <= arr(j, 1) Then '每个temp首先都和它上一个数比较
                arr(j + 1, 1) = arr(j, 1) 'temp小于该数则该数下沉一个位置留出插入空间
            Else
                Exit For
            End If
            arr(j + 1, 1) = temp 'temp大于时，temp插入被比较数的下一个数。
        Next j
    Next i
End Sub

'从上往下，将待插入值（外层循环的数值）插入到上层有序数列中。
'第一趟比较第1、2个数，得出含有两个数的有序数列
'注意：插入地址是内层循环结束后被比较数当前位置的后一个位置
'外层循环次数：需要选中N次temp与有序数列进行比较，故要循环N次。
'内层循环次数；因每次外层循环有序数列加1，故每次内层需比较次数至少为有序数列
'的数量为次数。
'小于有序数列最底端最大数时，最大数下沉占用temp的位置，由此可以为temp插入留出
'多余（1个）的空间


'插入排序在序列基本有序时，效率比较高
'插入排序在待排序个数较少时，效率比较高