Attribute VB_Name = "Module1"
Option Explicit


Function CreateCoinKit()

    Dim result As New Scripting.Dictionary
    
    Dim arr_temp_������_��������_�������(1 To 6) As String
    arr_temp_������_��������_�������(3) = vbYellow
    
    Dim i As Integer
    For i = Asc("A") To Asc("A") + 11
        arr_temp_������_��������_�������(6) = Chr(i)
        result.Add Chr(i), arr_temp_������_��������_�������
    Next i
    
    Set CreateCoinKit = result
End Function
