Attribute VB_Name = "Module2"
Sub test()
    Set Rng = Selection
    Rng.Copy Destination:=Worksheets("����3").Range("b3", "c4")
    Application.ScreenUpdating = False
    Worksheets("����3").Activate
    Range("a4").Select
    Worksheets("����3").Paste
    Worksheets("����3").Select
'    Debug.Print Rng.Cells(2, 1).Address
End Sub


Sub ����_����_���������_()
Stop
    Set Rng = Selection
    ������ = Range("a1").Interior.Color
    ������� = Range("b2").Interior.Color
    For Each cell In Rng
        If cell.Interior.Color = ������ Then
            cell.Interior.Color = �������
        Else
            cell.Interior.Color = ������
        End If
    Next cell
End Sub
Sub ������1()
Attribute ������1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������1 ������
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
