Attribute VB_Name = "Module1"
Option Explicit


Function CreateCoinKit()

    Dim result As New Scripting.Dictionary
    
    Dim arr_temp_пустая_сущность_монетки(1 To 6) As String
    arr_temp_пустая_сущность_монетки(3) = vbYellow
    
    Dim i As Integer
    For i = Asc("A") To Asc("A") + 11
        arr_temp_пустая_сущность_монетки(6) = Chr(i)
        result.Add Chr(i), arr_temp_пустая_сущность_монетки
    Next i
    
    Set CreateCoinKit = result
End Function
