VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Dim select_r As Integer

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Column = 2 And Target.Row > 11 Then
        Cells(2, 3).Value = Cells(Target.Row, 2).Value
        Cells(3, 3).Value = Cells(Target.Row, 1).Value
        Cells(4, 3).Value = Date
        Cells(7, 3).Value = 0
        Cells(8, 3).Value = ""
        Cells(9, 3).Value = Cells(Target.Row, 4).Value
        Range(Cells(Target.Row, 1), Cells(Target.Row, 11)).Interior.Color = 65535
        If Cells(6, 11).Value > 11 Then
            Range(Cells(Cells(6, 11).Value, 1), Cells(Cells(6, 11).Value, 11)).Interior.Color = RGB(255, 255, 255)
        End If
        Cells(6, 11).Value = Target.Row
    End If
End Sub

