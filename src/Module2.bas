Attribute VB_Name = "Module2"
Sub test()
    Set Rng = Selection
    Rng.Copy Destination:=Worksheets("Лист3").Range("b3", "c4")
    Application.ScreenUpdating = False
    Worksheets("Лист3").Activate
    Range("a4").Select
    Worksheets("Лист3").Paste
    Worksheets("Лист3").Select
'    Debug.Print Rng.Cells(2, 1).Address
End Sub


Sub меят_цвет_потрачено_()
Stop
    Set Rng = Selection
    черный = Range("a1").Interior.Color
    красный = Range("b2").Interior.Color
    For Each cell In Rng
        If cell.Interior.Color = черный Then
            cell.Interior.Color = красный
        Else
            cell.Interior.Color = черный
        End If
    Next cell
End Sub
Sub Макрос1()
Attribute Макрос1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Макрос1 Макрос
'

'
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
