' 以输入的 String 在末尾新建 sheet
Sub newSheets(Rng As String)

    Dim Sht As Worksheet
    Dim k As Integer

    For Each Sht In Sheets
        If Sht.Name = Rng Then
            k = 1
            Application.DisplayAlerts = False
            Sheets(Rng).Delete
            Application.DisplayAlerts = True
        End If
        Next

        If k = 0 Then
            Sheets.Add after:=Sheets(Sheets.Count)
            Sheets(Sheets.Count).Name = Rng
            Sheets("Sheet1").Select
        End If


End Sub

' 给 Combobox 添加数据，并且将最后一个数据设置成它的默认值
Private Sub addDataToCombobox()
    For i = 3 To 10

        UserForm1.ComboBox1.AddItem Sheet1.range("c" & i)
        Next

        UserForm1.ComboBox1.Value = Sheet1.range("c" & 10)

End Sub

Private Sub TextBox1_Change()
    If Len(Me.TextBox1.Value) >= 4 Then
        Me.ListBox1.AddItem Sheet1.Range("i" & i)
    End if
    Next
    If Me.ListBox1.ListCount >0 Then
        Me.ListBox.Visible = True
    Else
        Me.listbox.visible = false
    end if

    ' 使用 ListBox，参考 13课 1:00:12
End Sub

' 数组
sub learn_arr()
    dim arr(1 to 4)

    for i = 1 to 4
        arr(i) = range("b" & i + 1) * range("c" & i + 1)
        next

        range("h3") = application.worksheetFunction.max(arr)
end sub



' 数组 针对数据量不确定的情况
sub learn_arr_2()
    dim arr() ' 先不定义数组大小
    dim j, i As Integer

    j = range("a65536").end(xlup).Row + 1    ' +1是为了去表头
    redim arr(1 to j)    ' dim 不支持变量，而redim 可以，所以arr定义了两次

    for i = 1 to 4
        arr(i) = range("b" & i + 1) * range("c" & i + 1)
        next

        range("h3") = application.worksheetFunction.max(arr)
        range("h2") = range("a"& application.worksheetFunction.match(range("h3"), arr, 0) + 1)

        MsgBox Lbound(arr)   ' 下限
        MsgBox Ubound(arr)   ' 上限
end sub


' 数组应用： 排列组合暴力计算
' 已知一列80行的数字，区其中任意4个数字的和为124704
sub learn_arr_3()
    dim i, j, k, l As Integer
    dim arr()


    t = timer  ' 引入计时器，返回计算耗时

    arr = range("a1:a80")
    for i = 2 to 80
        for j = 2 to 80
            for k = 2 to 80
                for l = 2 to 80
                    if arr(i, 1) + arr(j, 1) + arr(k, 1) + arr(l, 1) = 124704 then
                        range("e3") = arr(i, 1)
                        range("f3") = arr(j, 1)
                        range("g3") = arr(k, 1)
                        range("h3") = arr(l, 1)
                        GoTO 100
                    end if
                    next
                    next
                    next
                    next

                    100

                    msgbox format(timer - t, "0.0000")
end sub


' 字典
Sub dic()

    Dim dic As New Dictionary

    dic.Add 1, "张三"
    dic.Add 2, "李四"

    range("a10") = dic(2)

End Sub

Sub dic_2()

    Dim dic As New Dictionary

    dic.Add "张三", 3000
    dic.Add "李四", 2000

    range("a10") = dic("李四")
    ' 这里的“李四”相当于KEY，指定他就能返回后面的值
End Sub

Sub dic_3()

    Dim dic As New Dictionary

    dic("李四") = 8000
    ' 赋值给 key "李四"

    range("a10") = dic("李四")
    ' 这里的“李四”相当于KEY，指定他就能返回后面的值
End Sub

Sub dic_4()

    Dim dic As New Dictionary

    for i = 2 to 5
        dic (range("d"*i).value) = range("e"& i).value  ' 这里的 value 一定不能省略

        range("a10") = dic(2)

End Sub


' 给下拉列表添加字典
Private Sub UserForm_Activate()
    Dim arr()
    Dim dic As New Dictionary

    arr = range("c3:d50")

    For i = LBound(arr()) To UBound(arr())

        dic(arr(i, 1)) = 1
        Next

        Me.ComboBox1.List = dic.Keys

End Sub