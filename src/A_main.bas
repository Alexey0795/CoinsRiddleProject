Attribute VB_Name = "A_main"
 Public ГЛОБ_МОНЕТЫ

Sub нажато_право()
    Call the_main("R")
End Sub
Sub нажато_лево()
    Call the_main("L")
End Sub
Sub нажато_стол()
    Call the_main("T")
End Sub
Sub нажато_ресет()
    Call рестарт
End Sub
Sub нажато_ответ()
Stop
    Set Rng = Selection
    If Rng.Cells.Count <> 1 Then Exit Sub
    
    Rng.Cut Destination:=Worksheets("Лист1").Range("q6")
    
    Call проверка_отета
'    Call рестарт
End Sub

Sub the_main(Optional меседж)
'    Stop
'    меседж = "R"
    the_area = Selection.Address 'сохраняем месеж
    Set селекшн = Selection
    
    Set ГЛОБ_МОНЕТЫ = CreateObject("Scripting.Dictionary")
    Set ГЛОБ_МОНЕТЫ = CreateCoinKit
    Call прочитать_всю_сцену
    
    'сначала проверить селекшин
    Set сло_набор = проверить_селекши(the_area)
    
    'заказать адрес
    сколько_монет = сло_набор.Count
    направление = меседж
    If направление <> "T" Then
        список_адресов = чаша(направление, сколько_монет)
    Else
        список_адресов = на_стол(сколько_монет)
    End If
    
    'применить адрес
    Set сло_изм = взаимодействие(список_адресов, сло_набор)
    If сло_изм.exists("проблема") Then Exit Sub
    Call соверш_перемещ_монет(сло_изм)
    
    If Application.Intersect(селекшн, Range("СТОЛ")) Is Nothing Then
        Call вызвать_гравитацию
    End If
    
End Sub

Sub рестарт()
'Stop
'    Call цвет_чаши(vbYellow, vbYellow)
    Worksheets("сцены").Range("весы_2").Copy Destination:=Worksheets("Лист1").Range("СЦ")
    Worksheets("Лист1").Range("a1", "d1").Clear
    Worksheets("Лист1").Range("p1", "p2").Clear
    Worksheets("Лист1").Range("q6").Clear
    
    Range("чаша_сц_Л_сре").Clear
    Range("чаша_сц_Л_сре").Interior.Color = vbWhite
    Range("чаша_сц_П_сре").Clear
    Range("чаша_сц_П_сре").Interior.Color = vbWhite
    
    Range("СТОЛ").Clear
    
    Worksheets("Лист2").Range("a29").Clear
    Set вирт_монеты = CreateCoinKit()
    Set ГЛОБ_МОНЕТЫ = прописать_монетки_на_столе(вирт_монеты)
    For Each Key In ГЛОБ_МОНЕТЫ
        temp_arr = ГЛОБ_МОНЕТЫ.Item(Key)
        рисовать (temp_arr)
        Call lag
    Next Key
    Call паршиватор
    Call конверт 'удалить из рестарта
End Sub

Function сорентироваться(буква, статус_весов)
    'вернет адрес текущего положения чаши про которую сспросили
    temp = буква & 5 + статус_весов '5 равновесие
    Select Case temp
        Case "L4"
            сорентироваться = Range("чаша_сц_Л_лег").Address
        Case "L5"
            сорентироваться = Range("чаша_сц_Л_сре").Address
        Case "L6"
            сорентироваться = Range("чаша_сц_Л_тяж").Address
        Case "R4"
            сорентироваться = Range("чаша_сц_П_тяж").Address
        Case "R5"
            сорентироваться = Range("чаша_сц_П_сре").Address
        Case "R6"
            сорентироваться = Range("чаша_сц_П_лег").Address
    End Select
    
End Function

Function хранит_адрес_чаши(код_запроса)
Stop
    Select Case код_запроса
    Case "L-"
        левая_чаша_минус = Range("e3", "g9").Address
        хранит_адрес_чаши = левая_чаша_минус
    Case "L0"
        левая_чаша_ноль = Range("e7", "g12").Address
        хранит_адрес_чаши = левая_чаша_ноль
    Case "L+"
        левая_чаша_плюс = Range("e11", "g16").Address
        хранит_адрес_чаши = левая_чаша_плюс
    Case "R-"
        правая_чаша_минус = Range("k3", "m16").Address
        хранит_адрес_чаши = правая_чаша_минус
    Case "R0"
        правая_чаша_ноль = Range("k7", "m12").Address
        хранит_адрес_чаши = правая_чаша_ноль
    Case "R+"
        правая_чаша_плюс = Range("k11", "m9").Address
        хранит_адрес_чаши = правая_чаша_плюс
    End Select
End Function



Function прописать_монетки_на_столе(монетки)
    'возвращает набор монеток с присвоеными адресами
    Set Rng = Range("СТОЛ")
    Set сло_набор_монеток = монетки
             
    'составляем перечень ячеек кандидатов
    Set сло_рабочие_ячейки_стола = CreateObject("Scripting.Dictionary")
    For Each cell In Rng
        If cell = "" Then
            n = n + 1
            Key = CStr(n)
            Item = cell.Address
            сло_рабочие_ячейки_стола.Add Key, Item
        End If
    Next cell
    
    'составляем перечень монеток кандидатов
    Set сло_монетки_кондидаты = CreateObject("Scripting.Dictionary")
    For Each Key In сло_набор_монеток
        Item = сло_набор_монеток.Item(Key)
        сло_монетки_кондидаты.Add Key, Item
    Next Key
    
    'присваиваем адреса набору монет
    For Each key_монеты In сло_монетки_кондидаты
        лимит_стола = сло_рабочие_ячейки_стола.Count
        лимит_монет = сло_монетки_кондидаты.Count
        temp = CStr(Int((лимит_стола * Rnd) + 1)) - 1 'порядковвый номер в словаре                   'на основе оставшихся мест
        
        key_стол = сло_рабочие_ячейки_стола.keys()(temp)
        temp_стол = сло_рабочие_ячейки_стола.items()(temp)
        
        arr_temp_монеты = сло_набор_монеток.Item(key_монеты)   'из основного массива
        arr_temp_монеты(2) = temp_стол
        сло_набор_монеток(key_монеты) = arr_temp_монеты           'сохр в осно массив
        
        сло_рабочие_ячейки_стола.Remove key_стол
        сло_монетки_кондидаты.Remove key_монеты
        
        aaa = сло_набор_монеток.Item(key_монеты)
        a1 = aaa(1)
        a2 = aaa(2)
        a3 = aaa(3)
        a4 = aaa(4)
        a5 = aaa(5)
        a6 = aaa(6)
    Next key_монеты
    
    Set прописать_монетки_на_столе = сло_набор_монеток
End Function

Sub рисовать(arr_temp)
    'наносит краску на одну ячейку
    'в итеме все неоходимое для покраски цвет буква адрес
    адрес = arr_temp(2)
    цвет = получить_краску(arr_temp(3))
    Worksheets("Лист1").Range(адрес).Interior.Color = цвет
    Worksheets("Лист1").Range(адрес) = arr_temp(6)
End Sub

Function получить_краску(слово_для_цвета)
    Select Case слово_для_цвета
        Case "эталон"
            получить_краску = vbBlue
        Case "L"
            получить_краску = vbGreen
        Case "R"
            получить_краску = vbMagenta
        Case "дефолт"
            получить_краску = vbYellow
        Case Else
            получить_краску = слово_для_цвета 'лень
    End Select
End Function

Sub lag()
   Dim a As Single
   a = Timer
   Do While Timer < a + 0.025
   Loop
End Sub

Sub прочитать_всю_сцену()
    Set temp_сло = прочитать_стол()
    Call записать_прочитаное(temp_сло)
    Set temp_сло = прочитать_чашу("L")
    Call записать_прочитаное(temp_сло)
    Set temp_сло = прочитать_чашу("R")
    Call записать_прочитаное(temp_сло)
End Sub

Sub записать_прочитаное(temp_сло)
    'если существует пустой ключ он не должен попасть в глобальную
    If temp_сло.exists("пусто") = True Then Exit Sub
'    буква = Left(конверт(), 1)
    For Each Key In temp_сло
        Item = temp_сло.Item(Key)
        If ГЛОБ_МОНЕТЫ.exists(Key) = False Then
            ГЛОБ_МОНЕТЫ.Add Key, Item 'если нет - создаем
        Else
            что_там = ГЛОБ_МОНЕТЫ.Item(Key)
            ГЛОБ_МОНЕТЫ.Item(Key) = Item 'если есть - перезаписываем
        End If
    Next Key
End Sub

Function прочитать_стол()
    adr = Range("СТОЛ").Address
    Dim temp_arr(1 To 6) 'для свойст ячейки
    Set сло_монеты_со_стола = CreateObject("Scripting.Dictionary")
    Set Rng = Range(adr)
    For Each cell In Rng
        If cell <> "" Then
            Key = cell.Value
            temp_arr(1) = "стол"
            temp_arr(2) = cell.Address
            temp_arr(3) = cell.Interior.Color
            temp_arr(4) = "неизв колонка" 'надо будет вызывать функции - нах надо
            temp_arr(5) = "неизв высота"
            temp_arr(6) = Key
            If сло_монеты_со_стола.exists(Key) = False Then
                сло_монеты_со_стола.Add Key, temp_arr
            Else
                сло_монеты_со_стола.Item(Key) = temp_arr
            End If
        End If
    Next cell
    Set прочитать_стол = сло_монеты_со_стола
End Function

Function прочитать_чашу(сторона)
    статус_весов = Range("m1")
    adr = сорентироваться(сторона, статус_весов)

    Set Rng = Range(adr)
    
    Dim arr_temp(1 To 6)
    Set сло_какаято_чаша = CreateObject("Scripting.Dictionary")
    For Each cell In Rng
        If cell <> "" Then
            Key = cell.Value
            arr_temp(1) = сторона
            arr_temp(2) = cell.Address
            arr_temp(3) = cell.Interior.Color
            arr_temp(4) = "колонка чаши"
            arr_temp(5) = "высота от чаши"
            arr_temp(6) = Key
            сло_какаято_чаша.Add Key, arr_temp
        End If
    Next cell
    
    найдено = сло_какаято_чаша.Count
    If найдено < 1 Then
        Key = "пусто"
        сло_какаято_чаша.Add Key, 0
    End If
    Set прочитать_чашу = сло_какаято_чаша
End Function
Function прочитать_статус_весов()
    прочитать_статус_весов = Range("m1")
End Function

Sub соверш_перемещ_монет(сло_реестр_изменений)
    'перемещение подразм что что-то где-то есть
    'надо стереть нарисовать и отметить в глобалке
    'получает словарь изменений в котором объект монеты,
    'на этот момент глобалка уже создана
    'задача функции отмыть краску и передать покращику итем

    For Each Key In сло_реестр_изменений
        'копируем старые даные, подтираем
        arr_temp = ГЛОБ_МОНЕТЫ.Item(Key)     'старый итем
        старый_адрес = arr_temp(2)           'старый адрес
        Range(старый_адрес).Clear
        Range(старый_адрес).Interior.Color = vbWhite
        arr_temp = сло_реестр_изменений.Item(Key) 'новый итем
        ГЛОБ_МОНЕТЫ.Item(Key) = arr_temp          'ред перем
        рисовать (arr_temp)   'только итем экземпляра
        
        Call lag
    Next Key
End Sub
Function проверить_селекши(the_area)
    'возращает словарь или ошибку с читыми объектами
    Set Rng = Range(the_area)
    Set сло_измен = CreateObject("Scripting.Dictionary")
    
    'если выбрано слишком много просто прекратить или вернуть сигнал
    If Selection.Count > 71 Then
        Debug.Print "71 ячеек"
        сло_измен.Add "проблема", 1010
        Set проверить_селекши = сло_измен
        Exit Function
    End If
    
    'работаем только с буквами проверяяем 12 букв
    Dim arr_temp(1 To 6)
    For Each cell In Rng
        If cell <> "" Then
            просмотр = cell.Value
            If Asc(просмотр) > 64 And Asc(просмотр) < 77 Then
                сколько_монет = сколько_монет + 1
                Key = cell.Value
                arr_temp(1) = направление
                arr_temp(2) = cell.Address
                arr_temp(3) = cell.Interior.Color
'                arr_temp(4) = "колонка чаши"
'                arr_temp(5) = "высота от чаши"
                arr_temp(6) = Key
                сло_измен.Add Key, arr_temp
            Else
                'вернуть сигнал
                Debug.Print "101"
                сло_измен.Add "проблема", 101
                Set проверить_селекши = сло_измен
                Exit Function
            End If
        End If
    Next cell
    
    'если букв в селекшине нету то подать сигнал
    If сло_измен.Count < 1 Then
        'вернуть сигнал
        Debug.Print "102"
        сло_измен.Add "проблема", 102
        Set проверить_селекши = сло_измен
        Exit Function
    End If
    
    Set проверить_селекши = сло_измен
End Function
Function взаимодействие(список_адресов, сло_измен)
    'подменяет адреса в итемах
    
    'если закончилось место на столе например
'    If список_адресов(UBound(список_адресов)) = "проблема" Then
'        'вернуть сигнал
'        Debug.Print "не влазит в стол"
'        сло_измен.Add "проблема", "не влазит в стол"
'        Set взаимодействие = сло_измен
'        Exit Function
'    End If
    
    'зашиваем полученые адреса по прямомму списку
    For Each Key In сло_измен
        q = q + 1
        Item = сло_измен.Item(Key)
        If список_адресов(q) = "проблема" Then
            Debug.Print "не влазит в стол"
            сло_измен.Add "проблема", "не влазит в стол"
            Set взаимодействие = сло_измен
            Exit Function
        End If
        Item(2) = список_адресов(q)
        сло_измен.Item(Key) = Item
    Next Key
    
    Set взаимодействие = сло_измен
End Function

Function чаша(направление, количество, Optional номер_фикс = "")
    'отдает список заселения
    'найти адрес
    adr = сорентироваться(направление, Range("m1"))
'    adr = хранит_адрес_чаши(направление & сост_весов)
    Set Rng = Range(adr)
    Set сло_костыль = CreateObject("Scripting.Dictionary")
    
    число_колон = Rng.Columns.Count
    число_высот = Rng.Rows.Count
    
    Dim список()
    
    For j = 1 To количество
        'зарандомить колонку
        If номер_фикс = "" Then
            номер_колонки = (Int((число_колон * Rnd) + 1))
        Else
            номер_колонки = номер_фикс
        End If
        
        'пролистать вверх
        'составить список
        For i = число_высот To 1 Step -1
            If Rng.Cells(i, номер_колонки) = "" Then
                некий_адрес = Rng.Cells(i, номер_колонки).Address
                If сло_костыль.exists(некий_адрес) Then
                    GoTo попробовать_еще_раз_выше
                End If
                n = n + 1
                ReDim Preserve список(1 To n)
                список(n) = Rng.Cells(i, номер_колонки).Address
                a_key = список(n)
                сло_костыль.Add a_key, 0 'занятые
'                Rng.Cells(i, номер_колонки).Interior.Color = vbRed
                GoTo след_адрес 'вот тут собака зарыта
            End If
попробовать_еще_раз_выше:
        Next i
        'если цикл закончился а свобод мест нету перейти на сосед колонку
        Select Case номер_колонки
            Case 1
                номер_колонки = 2
            Case 2
                номер_колонки = 3
            Case 3
                номер_колонки = 1
        End Select
        j = j - 1
        попытка = попытка + 1
        If попытка > 12 Then
            Stop
            Exit Function 'закладка
        End If
след_адрес:
    Next j
    
    чаша = список
End Function

Function на_стол(количество)
    'отдает список заселения для стола
  
'    adr = "СТОЛ"
    Set Rng = Range("СТОЛ")
    Set сло_костыль = CreateObject("Scripting.Dictionary")
        
    'составляем перечень ячеек кандидатов
    Set сло_рабочие_ячейки_стола = CreateObject("Scripting.Dictionary")
    For Each cell In Rng
        If cell = "" Then
            n = n + 1
            Key = CStr(n)
            Item = cell.Address
            сло_рабочие_ячейки_стола.Add Key, Item
        End If
    Next cell
        
    Dim список()
    
    For j = 1 To количество
        If сло_рабочие_ячейки_стола.Count = 0 Then
            q = q + 1
            ReDim Preserve список(1 To q)
            список(q) = "проблема"
            на_стол = список
            Exit Function
        End If
        'зарандомить индекс
        индекс = (Int((сло_рабочие_ячейки_стола.Count * Rnd) + 1))
        Key = сло_рабочие_ячейки_стола.keys()(индекс - 1)
        Item = сло_рабочие_ячейки_стола.Item(Key)
        q = q + 1
        ReDim Preserve список(1 To q)
        список(q) = Item
        сло_рабочие_ячейки_стола.Remove (Key)
    Next j
    
    на_стол = список
End Function

Sub паршиватор()
    число = ГЛОБ_МОНЕТЫ.Count
    овца = Int((число * Rnd) + 1)
    temp_arr = ГЛОБ_МОНЕТЫ.items()(овца - 1)
    направление_отличая = Int((2 * Rnd) + 1)
    Select Case направление_отличая
        Case 1
            temp_arr(1) = "-1"
        Case 2
            temp_arr(1) = "1"
    End Select
    Key = temp_arr(6)
    сегодня = CLng(Date)
    Range("P1") = (сегодня - Asc(Key) - 43000) * 10 + temp_arr(1)
End Sub

Function конверт()
    temp = Right(Worksheets("Лист1").Range("P1"), 1)
    If temp = 9 Then
        side = -1
        наклейка = 2
    Else
        If temp = 1 Then
            side = 1
            наклейка = 4
        End If
    End If
    st1 = Range("P1") - side
    st2 = st1 / 10
    st3 = st2 + 43000
    st4 = CLng(Date)
    конверт = Chr(CLng(Date) - 43000 - (Range("P1") - side) / 10) & наклейка
    Worksheets("Лист1").Range("P2") = конверт
End Function

Sub взвесить()
    'копирует на др лист затем высталяет всю сцену целиком
    
    Range("a1") = Range("a1") + 1
    If Range("a1") > 3 Then
        MsgBox "ПОТРАЧЕНО"
        Exit Sub
    End If
    
    Set ГЛОБ_МОНЕТЫ = CreateObject("Scripting.Dictionary")
    Set ГЛОБ_МОНЕТЫ = CreateCoinKit
    Call прочитать_всю_сцену
    чаша_А = сорентироваться("L", Range("m1"))
    чаша_Б = сорентироваться("R", Range("m1"))
    
    temp = конверт() 'вес и буква
    буква = Left(temp, 1)
    вес = Right(temp, 1)
    temp_arr = ГЛОБ_МОНЕТЫ.Item(буква)
    место = temp_arr(1)

    For Each Key In ГЛОБ_МОНЕТЫ
        temp_arr = ГЛОБ_МОНЕТЫ.Item(Key) 'считаем сколько монет на чашах
        If temp_arr(1) = "L" Then
            cnt_A = cnt_A + 1
        Else
            If temp_arr(1) = "R" Then
                cnt_B = cnt_B + 1
            End If
        End If
    Next Key
    
    If cnt_A = cnt_B Then
        flg_совершить_подсказку = 1
        If место = "L" Then
            flg = "A" & вес
            flg_запомнить_тяжесть = 1
        Else
            If место = "R" Then
                If вес = 2 Then
                    flg = "A" & вес + 2
                    flg_запомнить_тяжесть = 1
                End If
                If вес = 4 Then
                    flg = "A" & вес - 2
                    flg_запомнить_тяжесть = 1
                End If
             Else
                If место = "T" Then
                    
                End If
             End If
        End If
        
    Else
        flg_отменить_подсказку = 0
        If cnt_A > cnt_B Then
            flg = "A4"
        Else
            If cnt_A < cnt_B Then
                flg = "A2"
            End If
        End If
    End If
    
    'определили необход полож весов и собир сцену для копирования
    Select Case flg
        Case "A2"
            сцена_для_переноса = "весы_1"
            Worksheets("сцены").Range("чаша_1А").Clear
            Worksheets("сцены").Range("чаша_1А").Interior.Color = vbWhite
            Worksheets("сцены").Range("чаша_1Б").Clear
            Worksheets("сцены").Range("чаша_1Б").Interior.Color = vbWhite
            
            Worksheets("Лист1").Range(чаша_А).Copy _
                    Destination:=Worksheets("сцены").Range("чаша_1А")
            Worksheets("Лист1").Range(чаша_Б).Copy _
                    Destination:=Worksheets("сцены").Range("чаша_1Б")
        Case "A4"
            сцена_для_переноса = "весы_3"
            Worksheets("сцены").Range("чаша_3А").Clear
            Worksheets("сцены").Range("чаша_3Б").Clear
            Worksheets("сцены").Range("чаша_3А").Interior.Color = vbWhite
            Worksheets("сцены").Range("чаша_3Б").Interior.Color = vbWhite
                        
            Worksheets("Лист1").Range(чаша_А).Copy _
                    Destination:=Worksheets("сцены").Range("чаша_3А")
            Worksheets("Лист1").Range(чаша_Б).Copy _
                    Destination:=Worksheets("сцены").Range("чаша_3Б")
        Case Else
            сцена_для_переноса = "весы_2"
            Worksheets("сцены").Range("чаша_2А").Clear
            Worksheets("сцены").Range("чаша_2Б").Clear
            Worksheets("сцены").Range("чаша_2А").Interior.Color = vbWhite
            Worksheets("сцены").Range("чаша_2Б").Interior.Color = vbWhite
            
            Worksheets("Лист1").Range(чаша_А).Copy _
                    Destination:=Worksheets("сцены").Range("чаша_2А")
            Worksheets("Лист1").Range(чаша_Б).Copy _
                    Destination:=Worksheets("сцены").Range("чаша_2Б")
    End Select
    
    Worksheets("сцены").Range(сцена_для_переноса).Copy Destination:=Worksheets("Лист1").Range("СЦ")
    If flg_совершить_подсказку = 1 Then
        Call совершить_подсказку(flg_запомнить_тяжесть)
    End If
    
    Range("b1") = Range("c1")
    Range("c1") = Range("d1")
    Range("d1") = Range("m1")
End Sub

Sub вызвать_гравитацию()
    Set ГЛОБ_МОНЕТЫ = CreateObject("Scripting.Dictionary")
    Call CreateCoinKit
    Call прочитать_всю_сцену
    adr = сорентироваться("R", Range("m1"))
    Call гравитация(adr)
    adr = сорентироваться("L", Range("m1"))
    Call гравитация(adr)
End Sub

Sub гравитация(adr)
    Set Rng = Range(adr)
    n_cln = Rng.Columns.Count
    n_rw = Rng.Rows.Count
    
'    Dim temp_arr(1 To 6)
    Set сло_изм = CreateObject("Scripting.Dictionary")
    
    For q = n_cln To 1 Step -1
        For i = n_rw To 1 Step -1
'            rng.Cells(i, q).Select
            prosmotr = Rng.Cells(i, q)
        
            If Rng.Cells(i, q) = "" Then
                For R = i - 1 To 1 Step -1
                    If Rng.Cells(R, q) <> "" Then
                        Key = Rng.Cells(R, q).Value
                        Item = ГЛОБ_МОНЕТЫ.Item(Key)
                        место = Item(1)
                        temp_arr = чаша(место, 1, q)
                        Item(2) = temp_arr(1)
                        If сло_изм.exists(Key) = True Then
                            сло_изм.Item(Key) = Item
                        Else
                            сло_изм.Add Key, Item
                        End If
                        Call соверш_перемещ_монет(сло_изм)
                        сло_изм.RemoveAll
                    End If
                Next R
            End If
        
        
        Next i
    Next q
End Sub

Sub совершить_подсказку(flg_запомнить_тяжесть)

    положение_весов = Worksheets("лист1").Range("m1")
    
    adr_L = сорентироваться("L", Range("m1"))
    adr_R = сорентироваться("R", Range("m1"))
    
    цвет_лег = Worksheets("сцены").Range("a30").Interior.Color
    цвет_тяж = Worksheets("сцены").Range("a31").Interior.Color
        
    Select Case положение_весов
        Case 0
            For Each cell In Range(adr_L)
                If cell.Value <> "" Then
                    cell.Interior.Color = vbBlue
                End If
            Next cell
            For Each cell In Range(adr_R)
                If cell.Value <> "" Then
                    cell.Interior.Color = vbBlue
                End If
            Next cell
        Case 1
            Call пометить_монеты(adr_L, adr_R, цвет_тяж, цвет_лег)
            
            For Each cell In Range("СТОЛ")
                If cell.Value <> "" Then
                    cell.Interior.Color = vbBlue
                End If
            Next cell
            
        Case -1
            Call пометить_монеты(adr_L, adr_R, цвет_лег, цвет_тяж)
                                    
            For Each cell In Range("СТОЛ")
                If cell.Value <> "" Then
                    cell.Interior.Color = vbBlue
                End If
            Next cell
            
        End Select
        
'        if
'    If положение_весов = 0 Then
'            For Each cell In Range(adr_L)
'                If cell.Value <> "" Then
'
'                        cell.Interior.Color = vbBlue
'
'                End If
'            Next cell
'            For Each cell In Range(adr_R)
'                If cell.Value <> "" Then
'
'                        cell.Interior.Color = vbBlue
'
'                End If
'            Next cell
'        Else
'            If положение_весов = 1 Or положение_весов = -1 Then
'                For Each cell In Range(adr_L)
'                    If cell.Value <> "" Then
'                        If cell.Interior.Color = vbYellow Then
'                            cell.Interior.Color = цвет_L
'                        End If
'                    End If
'                Next cell
'                For Each cell In Range(adr_R)
'                    If cell.Value <> "" Then
'                        If cell.Interior.Color = vbYellow Then
'                            cell.Interior.Color = цвет_R
'                        End If
'                    End If
'                Next cell
'
'                For Each cell In Range("СТОЛ")
'                    If cell.Value <> "" Then
'                        cell.Interior.Color = vbBlue
'                    End If
'                Next cell
'            End If
'        End If
        
'    Call пометить_монеты(adr_L, adr_R, цвет_L, цвет_R)

    'чтобы после этого цвета чаш не смогли поменяться
    If flg_запомнить_тяжесть = 1 Then Worksheets("сцены").Range("a29") = 1
    End Sub

Sub пометить_монеты(adr_L, adr_R, цвет_L, цвет_R)
'    взвешивание = Range("a1")
        
    For Each cell In Range(adr_L)
        If cell.Value <> "" Then
            If cell.Interior.Color = vbYellow Then
                cell.Interior.Color = цвет_L
            End If
        End If
    Next cell
    
    For Each cell In Range(adr_R)
        If cell.Value <> "" Then
            If cell.Interior.Color = vbYellow Then
                cell.Interior.Color = цвет_R
            End If
        End If
    Next cell
    
End Sub

Sub заэтанолить()
    Set Rng = Selection
    For Each cell In Rng
        If cell.Value <> "" Then cell.Interior.Color = vbBlue
    Next cell
End Sub

'Sub цвет_чаши(левый_цвет, правый_цвет)
'    arr_имен_диап_чаш = Array("чаша_1А", _
'        "чаша_1Б", _
'        "чаша_2А", _
'        "чаша_2Б", _
'        "чаша_3А", _
'        "чаша_3Б")
'    For i = 0 To UBound(arr_имен_диап_чаш)
'        adr = Worksheets("сцены").Range(arr_имен_диап_чаш(i)).Address
'        Set Rng = Range(adr)
'
'        Set Rng = Rng.Cells(Rng.Cells.Count).Offset(1, -2)
''        Set Rng =
''        Rng.Select
'        Set Rng = Rng.Resize(1, 3)
''        Rng.Select
'        If i Mod 2 = 0 Then
'            Rng.Interior.Color = левый_цвет
'        Else
'            Rng.Interior.Color = правый_цвет
'        End If
'    Next i
'
'End Sub
Sub the_bot()

    adr_L = сорентироваться("L", Range("m1"))
    adr_R = сорентироваться("R", Range("m1"))
    
    num_L = Range(adr_L).Count
    num_R = Range(adr_R).Count
    num_T = Range("СТОЛ").Count
    
    Select Case Worksheets("Лист1").Range("levl")
        
        Case 3
            
        Case 2
        Case 1
        Case 0
        Case Else
    End Select
End Sub

Sub проверка_отета()
    a = Left(конверт, 1)
    b = Range("q6")
    
    If a = b Then
        Range("ОЧКИ") = Range("ОЧКИ") + 1
    Else
        Range("ОЧКИ").Clear
        Range("n1", "n24").Clear
        Exit Sub
    End If
    
    For i = 24 To 1 Step -1
        If Range("N" & i) = "" Then
            adr = Range("N" & i).Address
            GoTo прекратить_поиск
        End If
    Next i
    
прекратить_поиск:
    Range("q6").Cut Destination:=Worksheets("Лист1").Range(adr)
End Sub
