Attribute VB_Name = "b_bot"
Sub the_bot_main()

    цвет_лег = Worksheets("сцены").Range("a30").Interior.Color
    цвет_тяж = Worksheets("сцены").Range("a31").Interior.Color
    
    шаг = Range("a1")
    Select Case шаг
        Case ""
            GoTo первый_шаг:
        Case 1
            GoTo второй_шаг:
        
        Case Else
            Stop
    End Select
    Range("e29").Select
    
первый_шаг:
    '1-ый шаг
    откуда = Range("СТОЛ").Address
    Call выбрать_горсть2(откуда, 4, vbYellow)
    Call нажато_лево
    Call выбрать_горсть2(откуда, 4, vbYellow)
    Call нажато_право
    Call взвесить
    
второй_шаг:
    '2-ой шаг
    If Range("m1") = 0 Then
        откуда = Range("чаша_сц_Л_сре").Address
        Call выбрать_горсть2(откуда, 4, vbBlue)
        Call нажато_стол
        откуда = Range("чаша_сц_П_сре").Address
        Call выбрать_горсть2(откуда, 3, vbBlue)
        Call нажато_стол
        откуда = Range("СТОЛ").Address
        Call выбрать_горсть2(откуда, 2, vbYellow)
        Call нажато_лево
        Call выбрать_горсть2(откуда, 1, vbYellow)
        Call нажато_право
    Else
        If Range("m1") = -1 Then
            зел_чаша = "L"
            орж_чаша = "R"
        Else
            зел_чаша = "R"
            орж_чаша = "L"
        End If
        
        откуда = сорентироваться(зел_чаша, Range("m1"))
        Call выбрать_горсть2(откуда, 2, цвет_лег)
        If зел_чаша = "R" Then
            Call нажато_лево
            Call выбрать_горсть2(Range("СТОЛ").Address, 2, vbBlue)
            Call нажато_право
        Else
            Call нажато_право
            Call выбрать_горсть2(Range("СТОЛ").Address, 2, vbBlue)
            Call нажато_лево
        End If
        Call выбрать_горсть2(откуда, 1, цвет_лег)
        Call нажато_стол
                
        откуда = сорентироваться(орж_чаша, Range("m1"))
        Call выбрать_горсть2(откуда, 2, цвет_тяж)
        Call нажато_стол
        Call выбрать_горсть2(откуда, 1, цвет_тяж)
        If зел_чаша = "L" Then
            Call нажато_лево
            Call выбрать_горсть2(Range("СТОЛ").Address, 2, vbBlue)
            Call нажато_право
        Else
            Call нажато_право
            Call выбрать_горсть2(Range("СТОЛ").Address, 2, vbBlue)
            Call нажато_лево
        End If
        Call выбрать_горсть2(откуда, 1, цвет_лег)
        Call нажато_стол
    End If
    
    Call взвесить
    
    '3-ий шаг
    'посчитать цвета
    'прочитать сцену
    Set ГЛОБ_МОНЕТЫ = CreateObject("Scripting.Dictionary")
    Set ГЛОБ_МОНЕТЫ = CreateCoinKit
    Call прочитать_всю_сцену
    For Each Key In ГЛОБ_МОНЕТЫ
        Item = ГЛОБ_МОНЕТЫ.Item(Key)
        цвет = Item(3)
        Select Case цвет
            Case vbBlue
                cnt_син = cnt_син + 1
            Case vbYellow
                cnt_жел = cnt_син + 1
            Case цвет_лег
                cnt_зел = cnt_зел + 1
            Case цвет_тяж
                cnt_тяж = cnt_тяж + 1
        End Select
    Next Key
    
    
'    adr = сорентироваться("L", Range("m1"))
'    Set Rng = Range(adr)
'    For Each cell In Rng
'        цвет = cell.Interior.Color
'        Select Case цвет
'            Case vbBlue
'                cnt_син = cnt_син + 1
'            Case vbYellow
'                cnt_жел = cnt_син + 1
'            Case цвет_лег
'                cnt_зел = cnt_зел + 1
'            Case цвет_тяж
'                cnt_тяж = cnt_тяж + 1
'        End Select
'    Next cell
''____________________________________________
'    adr = сорентироваться("R", Range("m1"))
'    Set Rng = Range(adr)
'    For Each cell In Rng
'        цвет = cell.Interior.Color
'        Select Case цвет
'            Case vbBlue
'                cnt_син = cnt_син + 1
'            Case vbYellow
'                cnt_жел = cnt_син + 1
'            Case цвет_лег
'                cnt_зел = cnt_зел + 1
'            Case цвет_тяж
'                cnt_тяж = cnt_тяж + 1
'        End Select
'    Next cell
''____________________________________________
'    Set Rng = Range("СТОЛ")
'    For Each cell In Rng
'        цвет = cell.Interior.Color
'        Select Case цвет
'            Case vbBlue
'                cnt_син = cnt_син + 1
'            Case vbYellow
'                cnt_жел = cnt_син + 1
'            Case цвет_лег
'                cnt_зел = cnt_зел + 1
'            Case цвет_тяж
'                cnt_тяж = cnt_тяж + 1
'        End Select
'    Next cell
    
    сост_1 = Range("c1")
    a = DoEvents
    сост_2 = Range("m1").Value
'==================================================
    If cnt_син = 8 Then
    a = DoEvents
        If сост_1 = 1 And сост_2 = 1 Then
            Легко = Range("чаша_сц_П_лег").Address
            Тяжело = Range("чаша_сц_Л_тяж").Address
        Else
            Легко = Range("чаша_сц_Л_лег").Address
            Тяжело = Range("чаша_сц_П_тяж").Address
        End If
            
            откуда = Тяжело
            Call выбрать_горсть2(откуда, 1, цвет_тяж)
            орж_подозев = Selection
            Call выбрать_горсть2(откуда, 1, vbBlue)
            Call нажато_стол
            Call выбрать_горсть2(откуда, 1, цвет_лег)
            Call заэтанолить
            Call нажато_стол
            
            откуда = Легко
            Call выбрать_горсть2(откуда, 1, цвет_лег)
            зел_подозев = Selection
            Call выбрать_горсть2(откуда, 2, vbBlue)
            Call нажато_стол
            Call выбрать_горсть2(откуда, 1, цвет_тяж)
            Call заэтанолить
            Call нажато_стол
            
            Call взвесить
            a = DoEvents
            
            If Range("m1") = 0 Then
                Range("СТОЛ").Find(орж_подозев).Select
            Else
                Range("чаша_сц_П_лег").Find(зел_подозев).Select
            End If
            Call нажато_ответ
            Exit Sub
'        Else
'            Stop
'        End If
    Else
        If сост_1 = 0 And сост_2 = 0 Then
            Call выбрать_горсть2(Range("СТОЛ").Address, 1, vbYellow)
            Call нажато_ответ
            Exit Sub
        End If
        
        If сост_1 = 0 And сост_2 <> 0 Then
            If cnt_зел > cnt_тяж Then
                
            End If
            Call выбрать_горсть2(Range("СТОЛ").Address, 1, vbYellow)
            Call нажато_ответ
            Exit Sub
        End If
        Stop
    End If
'==================================================
    Stop
End Sub

Sub выбрать_горсть2(откуда, сколько, Optional список_цветов = "")
    'выделяет ячейки на листе
    Dim arr_выбраные()
    Set сло_выбраные = CreateObject("Scripting.Dictionary")
    
    Set Rng = Range(откуда)
    
    For Each cell In Rng
        
'        arr_item = ГЛОБ_МОНЕТЫ.Item(Key)
            
            If список_цветов <> "" Then
                If cell.Interior.Color = список_цветов Then
                    flg_проерка_цвета = 1
                Else
                    flg_проерка_цвета = 0
                End If
            Else
                flg_проерка_цвета = 1
            End If
            
            If flg_проерка_цвета = 1 Then
                cnt_набрано = cnt_набрано + 1
                
                ReDim Preserve arr_выбраные(1 To cnt_набрано)
                arr_выбраные(cnt_набрано) = cell.Address
                
                If cnt_набрано >= сколько Then
'                    Set выбрать_горсть = сло_выбраные
                    GoTo завершить_набор
                End If
            End If
    
    Next cell
завершить_набор:
    For i = 1 To UBound(arr_выбраные)
        строка = строка & arr_выбраные(i) & ","
    Next i
    строка = Left(строка, Len(строка) - 1)
    
    Worksheets("Лист1").Range(строка).Select
'    Set Rng = Range(откуда)
'    Set сло_выбраные = CreateObject("Scripting.Dictionary")
'    For Each Cell In Rng
'        While n < сколько
'            If Cell.Interior.Color = список_цветов Then
'                Key = Cell.Value
'                сло_выбраные.Add Key, 1
'            End If
'        Loop
'    Next Cell
'
'    Set выбрать_горсть = сло_выбраные
End Sub

