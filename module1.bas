Attribute VB_Name = "module1"
Sub copy_Click()
    YesOrNoAnswerToMessageBox = MsgBox("确认复制用户 的refer link: " + Cells(2, 3).Value, vbYesNo, "Hi, " + staff)

    If YesOrNoAnswerToMessageBox = vbNo Then
        Exit Sub
    End If
    Cells(9, 3).Copy
End Sub

Sub BalanceUpdate_Click()
    Sheets("customer_master").Select
    Dim ID As Integer
    Dim num As String
    Dim balance_new As Double
    Dim balance_old As Double
    Dim next_bill_new As Date
    Dim next_bill_old As Date
    Dim refer_new As Double
    Dim refer_old As Double
    Dim staff As String
            
    ID = Cells(3, 3).Value

    num = CStr(Cells(2, 3).Value)
    staff = CStr(Cells(5, 3).Value)
    'If (num <> Cells(ID + 11, 2).Value) Then
    '    MsgBox ("请确认输入ID和号码")
    '    Exit Sub
    'End If
    
    YesOrNoAnswerToMessageBox = MsgBox("确认为该用户 更新Balance: " + num, vbYesNo, "Hi, " + staff)

    If YesOrNoAnswerToMessageBox = vbNo Then
        Exit Sub
    End If
    balance_old = Cells(ID + 11, 6).Value
    balance_new = Cells(7, 3).Value
    If balance_new < 0 Then
        balance_new = 0
    End If

    
    refer_old = Cells(ID + 11, 11).Value
    refer_new = refer_old - balance_new + balance_old
    If refer_new < 0 Then
        refer_new = 0
    End If
    Cells(ID + 11, 6).Value = balance_new
    Cells(ID + 11, 11).Value = refer_new
    
                    Sheets("Update_history").Select
                    LastRow_his = Range("A1000000").End(xlUp).Row

                    Cells(LastRow_his + 1, 1).Value = ID
                    Cells(LastRow_his + 1, 2).Value = num
                    Cells(LastRow_his + 1, 3).Value = Date
                    Cells(LastRow_his + 1, 4).Value = staff
                    Cells(LastRow_his + 1, 5).Value = balance_new
                    Cells(LastRow_his + 1, 6).Value = balance_old
                    
                    Cells(LastRow_his + 1, 7).Value = ""
                    Cells(LastRow_his + 1, 8).Value = ""
                    Cells(LastRow_his + 1, 9).Value = refer_new
                    Cells(LastRow_his + 1, 10).Value = refer_old
                    Cells(LastRow_his + 1, 11).Value = "Bal"
    Sheets("customer_master").Select
    MsgBox ("已成功为该用户 更新Balance: " + num + "   Refer链接已复制")
    Cells(ID + 11, 4).Copy

End Sub
Sub BillDateUpdate_Click()
    Sheets("customer_master").Select
            Dim ID As Integer
            Dim num As String
            Dim billdate_new As Date
            Dim billdate_old As Date

            Dim staff As String
            
    ID = Cells(3, 3).Value

    num = CStr(Cells(2, 3).Value)
    staff = CStr(Cells(5, 3).Value)
    'If (num <> Cells(ID + 11, 2).Value) Then
    '    MsgBox ("请确认输入ID和号码")
    '    Exit Sub
    'End If
    
    YesOrNoAnswerToMessageBox = MsgBox("确认为该用户 更新Bill Date: " + num, vbYesNo, "Hi, " + staff)

    If YesOrNoAnswerToMessageBox = vbNo Then
        Exit Sub
    End If
    billdate_old = Cells(ID + 11, 9).Value
    billdate_new = Cells(8, 3).Value
    Cells(ID + 11, 9).Value = billdate_new
                    Sheets("Update_history").Select
                    LastRow_his = Range("A1000000").End(xlUp).Row

                    Cells(LastRow_his + 1, 1).Value = ID
                    Cells(LastRow_his + 1, 2).Value = num
                    Cells(LastRow_his + 1, 3).Value = Date
                    Cells(LastRow_his + 1, 4).Value = staff
                    Cells(LastRow_his + 1, 5).Value = ""
                    Cells(LastRow_his + 1, 6).Value = ""
                    
                    Cells(LastRow_his + 1, 7).Value = billdate_new
                    Cells(LastRow_his + 1, 8).Value = billdate_old
                    Cells(LastRow_his + 1, 9).Value = ""
                    Cells(LastRow_his + 1, 10).Value = ""
                    Cells(LastRow_his + 1, 11).Value = "date"
    Sheets("customer_master").Select
    
End Sub

Sub newCustomer_Click()
    Sheets("customer_master").Select
    
    If Sheets("customer_master").AutoFilterMode Then
        Sheets("customer_master").AutoFilterMode = False
    End If
    YesOrNoAnswerToMessageBox = MsgBox("确认要添加新客户: " + Cells(3, 6).Value, vbYesNo, "确认信息")

    If YesOrNoAnswerToMessageBox = vbNo Then
        Exit Sub
    End If

    LastRow = Range("A1000000").End(xlUp).Row
    
    
    Range("$A$11:$L$" & LastRow + 1).AutoFilter
    Worksheets("customer_master").AutoFilter.Sort.SortFields.Add2 _
        Key:=Range("$A$11:$A$" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("customer_master").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    If (LastRow = 11) Then
        Cells(LastRow + 1, 1).Value = 1
    Else
        Cells(LastRow + 1, 1).Value = Cells(LastRow, 1).Value + 1
    End If
    
    Cells(LastRow + 1, 2).Value = Cells(3, 6).Value
    Cells(LastRow + 1, 3).Value = Cells(4, 6).Value
    Cells(LastRow + 1, 4).Value = Cells(6, 6).Value
    Cells(LastRow + 1, 5).Value = Cells(5, 6).Value
    Cells(LastRow + 1, 6).Value = Cells(7, 6).Value
    Cells(LastRow + 1, 7).Value = Cells(8, 6).Value
    Cells(LastRow + 1, 8).Value = Cells(9, 6).Value
    Cells(LastRow + 1, 9).Value = CDate(DateAdd("d", 28, Format(Cells(9, 6).Value, "dd/mm/yyyy")))
    Dim Month As Integer
    Month = Cells(LastRow + 1, 7).Value / Cells(LastRow + 1, 5).Value
    Cells(LastRow + 1, 10).Value = CDate(DateAdd("d", 28 * Month, Format(Cells(9, 6).Value, "dd/mm/yyyy")))
    Cells(LastRow + 1, 11).Value = Cells(LastRow + 1, 7).Value - Cells(LastRow + 1, 6).Value - Cells(LastRow + 1, 5).Value
  
    LastRow = LastRow + 1
    If Worksheets("customer_master").CheckBox2.Value = True Then
    Range("$A$11:$L$" & LastRow + 1).AutoFilter field:=12, Criteria1:="<>Yes"
    Worksheets("customer_master").AutoFilter.Sort.SortFields.Add2 _
        Key:=Range("$A$11:$A$" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("customer_master").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    End If
    MsgBox ("已成功添加新客户: " + Cells(3, 6).Value)

End Sub

Sub filter_Click()
    Call refresh_Click
    Sheets("customer_master").Select

    If Sheets("customer_master").AutoFilterMode Then
        Sheets("customer_master").AutoFilterMode = False
    End If

            
    LastRow = Range("A1000000").End(xlUp).Row
    Worksheets("customer_master").Range("$A$11:$L$" & LastRow + 1).AutoFilter
    Worksheets("customer_master").AutoFilter.Sort.SortFields.Add2 _
        Key:=Range("$A$11:$A$" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("customer_master").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Worksheets("customer_master").AutoFilter.Sort.SortFields.Clear
    
    If Worksheets("customer_master").CheckBox2.Value = True Then
        Range("$A$11:$L$" & LastRow).AutoFilter field:=12, Criteria1:="<>Yes", field:=11, Criteria1:="<>0"

        If Worksheets("customer_master").OptionButton1.Value = True Then
            Worksheets("customer_master").AutoFilter.Sort.SortFields.Add2 _
                Key:=Range("$K$11:$K$" & LastRow), SortOn:=xlSortOnValues, Order:=xlDescending, _
                DataOption:=xlSortNormal
            With Worksheets("customer_master").AutoFilter.Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End If

        If Worksheets("customer_master").OptionButton2.Value = True Then
            
            Worksheets("customer_master").AutoFilter.Sort.SortFields.Add2 _
            Key:=Range("$I$11:$I$" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
            With Worksheets("customer_master").AutoFilter.Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End If
    Else
       Range("$A$11:$L$" & LastRow).AutoFilter , field:=11, Criteria1:="<>0"
        
        If Worksheets("customer_master").OptionButton1.Value = True Then
            Worksheets("customer_master").AutoFilter.Sort.SortFields.Add2 _
                Key:=Range("$K$11:$K$" & LastRow), SortOn:=xlSortOnValues, Order:=xlDescending, _
                DataOption:=xlSortNormal
            With Worksheets("customer_master").AutoFilter.Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End If
    
        If Worksheets("customer_master").OptionButton2.Value = True Then
            
            
            Worksheets("customer_master").AutoFilter.Sort.SortFields.Add2 _
            Key:=Range("$I$11:$I$" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, _
            DataOption:=xlSortNormal
            With Worksheets("customer_master").AutoFilter.Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End If

    End If


End Sub

Sub clear_Click()
    Sheets("customer_master").Select
    Worksheets("customer_master").CheckBox2.Value = False
    If Sheets("customer_master").AutoFilterMode Then
        Sheets("customer_master").AutoFilterMode = False
    End If
    LastRow = Range("A1000000").End(xlUp).Row
  
    Range("$A$11:$L$" & LastRow + 1).AutoFilter
    Worksheets("customer_master").AutoFilter.Sort.SortFields.Clear
    Worksheets("customer_master").AutoFilter.Sort.SortFields.Add2 _
        Key:=Range("$A$11:$A$" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("customer_master").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub refresh_Click()
    Sheets("customer_master").Select
    Dim ID As Integer
    Dim num As String
    Dim balance_new As Double
    Dim balance_old As Double
    Dim next_bill_new As Date
    Dim next_bill_old As Date
    Dim refer_new As Double
    Dim refer_old As Double
    Dim Month As Integer
    Dim C_date, A_date, E_date, B_date As Date
    
    C_date = CDate(Format(Date, "dd/mm/yyyy"))
    Cells(1, 8).Value = C_date
    
    If Sheets("customer_master").AutoFilterMode Then
        Sheets("customer_master").AutoFilterMode = False
    End If

    LastRow = Range("A1000000").End(xlUp).Row
    Worksheets("customer_master").CheckBox2.Value = False
    Range("$A$11:$L$" & LastRow + 1).AutoFilter
    Worksheets("customer_master").AutoFilter.Sort.SortFields.Add2 _
        Key:=Range("$A$11:$A$" & LastRow), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("customer_master").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
     For i = 12 To LastRow
             If Cells(i, 11).Value = 0 Then
                 Cells(i, 12).Value = "Yes"
             End If
                A_date = CDate(Format(Cells(i, 8).Value, "dd/mm/yyyy"))
                B_date = CDate(Format(Cells(i, 9).Value, "dd/mm/yyyy"))
                E_date = CDate(Format(Cells(i, 10).Value, "dd/mm/yyyy"))


                If (B_date < C_date) And (C_date <= E_date) And (Cells(i, 12).Value <> "Yes") Then
                    ID = Cells(i, 1).Value
                    num = Cells(i, 2).Value
                    balance_old = Cells(i, 6).Value
                    next_bill_old = Cells(i, 9).Value
                    'refer_old = Cells(i, 11).Value
                    
                    Month = DateDiff("d", B_date, C_date) \ 28 + 1
                    
                    Cells(i, 9).Value = CDate(Format(DateAdd("d", 28 * Month, B_date), "dd/mm/yyyy"))
                    'Cells(i, 11).Value = Cells(i, 11).Value - Cells(i, 5).Value * Month
                    'If Cells(i, 11).Value < 0 Then
                    '    Cells(i, 11).Value = 0
                    'End If
                    Cells(i, 6).Value = Cells(i, 6).Value - Cells(i, 5).Value * Month
                    If Cells(i, 6).Value < 0 Then
                        Cells(i, 6).Value = 0
                    End If
                    
                    balance_new = Cells(i, 6).Value
                    next_bill_new = Cells(i, 9).Value
                    'refer_new = Cells(i, 11).Value
                    
                    Sheets("Update_history").Select
                    LastRow_his = Range("A1000000").End(xlUp).Row

                    Cells(LastRow_his + 1, 1).Value = ID
                    Cells(LastRow_his + 1, 2).Value = num
                    Cells(LastRow_his + 1, 3).Value = Date
                    Cells(LastRow_his + 1, 4).Value = "SYS"
                    Cells(LastRow_his + 1, 5).Value = balance_new
                    Cells(LastRow_his + 1, 6).Value = balance_old
                    
                                      
                    Cells(LastRow_his + 1, 7).Value = next_bill_new
                    Cells(LastRow_his + 1, 8).Value = next_bill_old
                    'Cells(LastRow_his + 1, 9).Value = refer_new
                    'Cells(LastRow_his + 1, 10).Value = refer_old
                    Cells(LastRow_his + 1, 11).Value = "refresh"
                    Sheets("customer_master").Select
                End If
            Next
End Sub

Sub detail_Click()
    Sheets("customer_master").Select
    Dim ID, c_paid, c_unpaid As Integer
    Dim num, plan As String
    Dim Active_date, end_date, B_date As Date

            
    ID = Cells(3, 3).Value
    num = CStr(Cells(2, 3).Value)
    plan = Cells(ID + 11, 5).Value
    Active_date = Cells(ID + 11, 8).Value
    end_date = Cells(ID + 11, 10).Value
    
    Dim Month As Integer
    Month = Cells(ID + 11, 7).Value / Cells(ID + 11, 5).Value
    
    Sheets("Bill_date_detail").Select
    Range("B:B").Value = ""
    Range("C:C").Value = ""
    Cells(11, 2).Value = ID
    Cells(12, 2).Value = num
    Cells(13, 2).Value = plan

    Cells(16, 2).Value = Active_date
    Cells(17, 2).Value = end_date
    c_paid = 0
    c_unpaid = 0
    For i = 0 To Month - 1
        B_date = CDate(Active_date + i * 28)
        Cells(18 + i, 2).Value = B_date
        If B_date <= Date Then
            c_paid = c_paid + 1
            Cells(18 + i, 3).Value = "ü"
        Else
            c_unpaid = c_unpaid + 1
        End If
    Next
    Cells(14, 2).Value = c_paid
    Cells(15, 2).Value = c_unpaid

End Sub


