Attribute VB_Name = "A_main"
 Public ����_������

Sub ������_�����()
    Call the_main("R")
End Sub
Sub ������_����()
    Call the_main("L")
End Sub
Sub ������_����()
    Call the_main("T")
End Sub
Sub ������_�����()
    Call �������
End Sub
Sub ������_�����()
Stop
    Set Rng = Selection
    If Rng.Cells.Count <> 1 Then Exit Sub
    
    Rng.Cut Destination:=Worksheets("����1").Range("q6")
    
    Call ��������_�����
'    Call �������
End Sub

Sub the_main(Optional ������)
'    Stop
'    ������ = "R"
    the_area = Selection.Address '��������� �����
    Set ������� = Selection
    
    Set ����_������ = CreateObject("Scripting.Dictionary")
    Set ����_������ = CreateCoinKit
    Call ���������_���_�����
    
    '������� ��������� ��������
    Set ���_����� = ���������_�������(the_area)
    
    '�������� �����
    �������_����� = ���_�����.Count
    ����������� = ������
    If ����������� <> "T" Then
        ������_������� = ����(�����������, �������_�����)
    Else
        ������_������� = ��_����(�������_�����)
    End If
    
    '��������� �����
    Set ���_��� = ��������������(������_�������, ���_�����)
    If ���_���.exists("��������") Then Exit Sub
    Call ������_�������_�����(���_���)
    
    If Application.Intersect(�������, Range("����")) Is Nothing Then
        Call �������_����������
    End If
    
End Sub

Sub �������()
'Stop
'    Call ����_����(vbYellow, vbYellow)
    Worksheets("�����").Range("����_2").Copy Destination:=Worksheets("����1").Range("��")
    Worksheets("����1").Range("a1", "d1").Clear
    Worksheets("����1").Range("p1", "p2").Clear
    Worksheets("����1").Range("q6").Clear
    
    Range("����_��_�_���").Clear
    Range("����_��_�_���").Interior.Color = vbWhite
    Range("����_��_�_���").Clear
    Range("����_��_�_���").Interior.Color = vbWhite
    
    Range("����").Clear
    
    Worksheets("����2").Range("a29").Clear
    Set ����_������ = CreateCoinKit()
    Set ����_������ = ���������_�������_��_�����(����_������)
    For Each Key In ����_������
        temp_arr = ����_������.Item(Key)
        �������� (temp_arr)
        Call lag
    Next Key
    Call ����������
    Call ������� '������� �� ��������
End Sub

Function ���������������(�����, ������_�����)
    '������ ����� �������� ��������� ���� ��� ������� ���������
    temp = ����� & 5 + ������_����� '5 ����������
    Select Case temp
        Case "L4"
            ��������������� = Range("����_��_�_���").Address
        Case "L5"
            ��������������� = Range("����_��_�_���").Address
        Case "L6"
            ��������������� = Range("����_��_�_���").Address
        Case "R4"
            ��������������� = Range("����_��_�_���").Address
        Case "R5"
            ��������������� = Range("����_��_�_���").Address
        Case "R6"
            ��������������� = Range("����_��_�_���").Address
    End Select
    
End Function

Function ������_�����_����(���_�������)
Stop
    Select Case ���_�������
    Case "L-"
        �����_����_����� = Range("e3", "g9").Address
        ������_�����_���� = �����_����_�����
    Case "L0"
        �����_����_���� = Range("e7", "g12").Address
        ������_�����_���� = �����_����_����
    Case "L+"
        �����_����_���� = Range("e11", "g16").Address
        ������_�����_���� = �����_����_����
    Case "R-"
        ������_����_����� = Range("k3", "m16").Address
        ������_�����_���� = ������_����_�����
    Case "R0"
        ������_����_���� = Range("k7", "m12").Address
        ������_�����_���� = ������_����_����
    Case "R+"
        ������_����_���� = Range("k11", "m9").Address
        ������_�����_���� = ������_����_����
    End Select
End Function



Function ���������_�������_��_�����(�������)
    '���������� ����� ������� � ����������� ��������
    Set Rng = Range("����")
    Set ���_�����_������� = �������
             
    '���������� �������� ����� ����������
    Set ���_�������_������_����� = CreateObject("Scripting.Dictionary")
    For Each cell In Rng
        If cell = "" Then
            n = n + 1
            Key = CStr(n)
            Item = cell.Address
            ���_�������_������_�����.Add Key, Item
        End If
    Next cell
    
    '���������� �������� ������� ����������
    Set ���_�������_��������� = CreateObject("Scripting.Dictionary")
    For Each Key In ���_�����_�������
        Item = ���_�����_�������.Item(Key)
        ���_�������_���������.Add Key, Item
    Next Key
    
    '����������� ������ ������ �����
    For Each key_������ In ���_�������_���������
        �����_����� = ���_�������_������_�����.Count
        �����_����� = ���_�������_���������.Count
        temp = CStr(Int((�����_����� * Rnd) + 1)) - 1 '����������� ����� � �������                   '�� ������ ���������� ����
        
        key_���� = ���_�������_������_�����.keys()(temp)
        temp_���� = ���_�������_������_�����.items()(temp)
        
        arr_temp_������ = ���_�����_�������.Item(key_������)   '�� ��������� �������
        arr_temp_������(2) = temp_����
        ���_�����_�������(key_������) = arr_temp_������           '���� � ���� ������
        
        ���_�������_������_�����.Remove key_����
        ���_�������_���������.Remove key_������
        
        aaa = ���_�����_�������.Item(key_������)
        a1 = aaa(1)
        a2 = aaa(2)
        a3 = aaa(3)
        a4 = aaa(4)
        a5 = aaa(5)
        a6 = aaa(6)
    Next key_������
    
    Set ���������_�������_��_����� = ���_�����_�������
End Function

Sub ��������(arr_temp)
    '������� ������ �� ���� ������
    '� ����� ��� ���������� ��� �������� ���� ����� �����
    ����� = arr_temp(2)
    ���� = ��������_������(arr_temp(3))
    Worksheets("����1").Range(�����).Interior.Color = ����
    Worksheets("����1").Range(�����) = arr_temp(6)
End Sub

Function ��������_������(�����_���_�����)
    Select Case �����_���_�����
        Case "������"
            ��������_������ = vbBlue
        Case "L"
            ��������_������ = vbGreen
        Case "R"
            ��������_������ = vbMagenta
        Case "������"
            ��������_������ = vbYellow
        Case Else
            ��������_������ = �����_���_����� '����
    End Select
End Function

Sub lag()
   Dim a As Single
   a = Timer
   Do While Timer < a + 0.025
   Loop
End Sub

Sub ���������_���_�����()
    Set temp_��� = ���������_����()
    Call ��������_����������(temp_���)
    Set temp_��� = ���������_����("L")
    Call ��������_����������(temp_���)
    Set temp_��� = ���������_����("R")
    Call ��������_����������(temp_���)
End Sub

Sub ��������_����������(temp_���)
    '���� ���������� ������ ���� �� �� ������ ������� � ����������
    If temp_���.exists("�����") = True Then Exit Sub
'    ����� = Left(�������(), 1)
    For Each Key In temp_���
        Item = temp_���.Item(Key)
        If ����_������.exists(Key) = False Then
            ����_������.Add Key, Item '���� ��� - �������
        Else
            ���_��� = ����_������.Item(Key)
            ����_������.Item(Key) = Item '���� ���� - ��������������
        End If
    Next Key
End Sub

Function ���������_����()
    adr = Range("����").Address
    Dim temp_arr(1 To 6) '��� ������ ������
    Set ���_������_��_����� = CreateObject("Scripting.Dictionary")
    Set Rng = Range(adr)
    For Each cell In Rng
        If cell <> "" Then
            Key = cell.Value
            temp_arr(1) = "����"
            temp_arr(2) = cell.Address
            temp_arr(3) = cell.Interior.Color
            temp_arr(4) = "����� �������" '���� ����� �������� ������� - ��� ����
            temp_arr(5) = "����� ������"
            temp_arr(6) = Key
            If ���_������_��_�����.exists(Key) = False Then
                ���_������_��_�����.Add Key, temp_arr
            Else
                ���_������_��_�����.Item(Key) = temp_arr
            End If
        End If
    Next cell
    Set ���������_���� = ���_������_��_�����
End Function

Function ���������_����(�������)
    ������_����� = Range("m1")
    adr = ���������������(�������, ������_�����)

    Set Rng = Range(adr)
    
    Dim arr_temp(1 To 6)
    Set ���_�������_���� = CreateObject("Scripting.Dictionary")
    For Each cell In Rng
        If cell <> "" Then
            Key = cell.Value
            arr_temp(1) = �������
            arr_temp(2) = cell.Address
            arr_temp(3) = cell.Interior.Color
            arr_temp(4) = "������� ����"
            arr_temp(5) = "������ �� ����"
            arr_temp(6) = Key
            ���_�������_����.Add Key, arr_temp
        End If
    Next cell
    
    ������� = ���_�������_����.Count
    If ������� < 1 Then
        Key = "�����"
        ���_�������_����.Add Key, 0
    End If
    Set ���������_���� = ���_�������_����
End Function
Function ���������_������_�����()
    ���������_������_����� = Range("m1")
End Function

Sub ������_�������_�����(���_������_���������)
    '����������� ������� ��� ���-�� ���-�� ����
    '���� ������� ���������� � �������� � ��������
    '�������� ������� ��������� � ������� ������ ������,
    '�� ���� ������ �������� ��� �������
    '������ ������� ������ ������ � �������� ��������� ����

    For Each Key In ���_������_���������
        '�������� ������ �����, ���������
        arr_temp = ����_������.Item(Key)     '������ ����
        ������_����� = arr_temp(2)           '������ �����
        Range(������_�����).Clear
        Range(������_�����).Interior.Color = vbWhite
        arr_temp = ���_������_���������.Item(Key) '����� ����
        ����_������.Item(Key) = arr_temp          '��� �����
        �������� (arr_temp)   '������ ���� ����������
        
        Call lag
    Next Key
End Sub
Function ���������_�������(the_area)
    '��������� ������� ��� ������ � ������ ���������
    Set Rng = Range(the_area)
    Set ���_����� = CreateObject("Scripting.Dictionary")
    
    '���� ������� ������� ����� ������ ���������� ��� ������� ������
    If Selection.Count > 71 Then
        Debug.Print "71 �����"
        ���_�����.Add "��������", 1010
        Set ���������_������� = ���_�����
        Exit Function
    End If
    
    '�������� ������ � ������� ���������� 12 ����
    Dim arr_temp(1 To 6)
    For Each cell In Rng
        If cell <> "" Then
            �������� = cell.Value
            If Asc(��������) > 64 And Asc(��������) < 77 Then
                �������_����� = �������_����� + 1
                Key = cell.Value
                arr_temp(1) = �����������
                arr_temp(2) = cell.Address
                arr_temp(3) = cell.Interior.Color
'                arr_temp(4) = "������� ����"
'                arr_temp(5) = "������ �� ����"
                arr_temp(6) = Key
                ���_�����.Add Key, arr_temp
            Else
                '������� ������
                Debug.Print "101"
                ���_�����.Add "��������", 101
                Set ���������_������� = ���_�����
                Exit Function
            End If
        End If
    Next cell
    
    '���� ���� � ��������� ���� �� ������ ������
    If ���_�����.Count < 1 Then
        '������� ������
        Debug.Print "102"
        ���_�����.Add "��������", 102
        Set ���������_������� = ���_�����
        Exit Function
    End If
    
    Set ���������_������� = ���_�����
End Function
Function ��������������(������_�������, ���_�����)
    '��������� ������ � ������
    
    '���� ����������� ����� �� ����� ��������
'    If ������_�������(UBound(������_�������)) = "��������" Then
'        '������� ������
'        Debug.Print "�� ������ � ����"
'        ���_�����.Add "��������", "�� ������ � ����"
'        Set �������������� = ���_�����
'        Exit Function
'    End If
    
    '�������� ��������� ������ �� �������� ������
    For Each Key In ���_�����
        q = q + 1
        Item = ���_�����.Item(Key)
        If ������_�������(q) = "��������" Then
            Debug.Print "�� ������ � ����"
            ���_�����.Add "��������", "�� ������ � ����"
            Set �������������� = ���_�����
            Exit Function
        End If
        Item(2) = ������_�������(q)
        ���_�����.Item(Key) = Item
    Next Key
    
    Set �������������� = ���_�����
End Function

Function ����(�����������, ����������, Optional �����_���� = "")
    '������ ������ ���������
    '����� �����
    adr = ���������������(�����������, Range("m1"))
'    adr = ������_�����_����(����������� & ����_�����)
    Set Rng = Range(adr)
    Set ���_������� = CreateObject("Scripting.Dictionary")
    
    �����_����� = Rng.Columns.Count
    �����_����� = Rng.Rows.Count
    
    Dim ������()
    
    For j = 1 To ����������
        '����������� �������
        If �����_���� = "" Then
            �����_������� = (Int((�����_����� * Rnd) + 1))
        Else
            �����_������� = �����_����
        End If
        
        '���������� �����
        '��������� ������
        For i = �����_����� To 1 Step -1
            If Rng.Cells(i, �����_�������) = "" Then
                �����_����� = Rng.Cells(i, �����_�������).Address
                If ���_�������.exists(�����_�����) Then
                    GoTo �����������_���_���_����
                End If
                n = n + 1
                ReDim Preserve ������(1 To n)
                ������(n) = Rng.Cells(i, �����_�������).Address
                a_key = ������(n)
                ���_�������.Add a_key, 0 '�������
'                Rng.Cells(i, �����_�������).Interior.Color = vbRed
                GoTo ����_����� '��� ��� ������ ������
            End If
�����������_���_���_����:
        Next i
        '���� ���� ���������� � ������ ���� ���� ������� �� ����� �������
        Select Case �����_�������
            Case 1
                �����_������� = 2
            Case 2
                �����_������� = 3
            Case 3
                �����_������� = 1
        End Select
        j = j - 1
        ������� = ������� + 1
        If ������� > 12 Then
            Stop
            Exit Function '��������
        End If
����_�����:
    Next j
    
    ���� = ������
End Function

Function ��_����(����������)
    '������ ������ ��������� ��� �����
  
'    adr = "����"
    Set Rng = Range("����")
    Set ���_������� = CreateObject("Scripting.Dictionary")
        
    '���������� �������� ����� ����������
    Set ���_�������_������_����� = CreateObject("Scripting.Dictionary")
    For Each cell In Rng
        If cell = "" Then
            n = n + 1
            Key = CStr(n)
            Item = cell.Address
            ���_�������_������_�����.Add Key, Item
        End If
    Next cell
        
    Dim ������()
    
    For j = 1 To ����������
        If ���_�������_������_�����.Count = 0 Then
            q = q + 1
            ReDim Preserve ������(1 To q)
            ������(q) = "��������"
            ��_���� = ������
            Exit Function
        End If
        '����������� ������
        ������ = (Int((���_�������_������_�����.Count * Rnd) + 1))
        Key = ���_�������_������_�����.keys()(������ - 1)
        Item = ���_�������_������_�����.Item(Key)
        q = q + 1
        ReDim Preserve ������(1 To q)
        ������(q) = Item
        ���_�������_������_�����.Remove (Key)
    Next j
    
    ��_���� = ������
End Function

Sub ����������()
    ����� = ����_������.Count
    ���� = Int((����� * Rnd) + 1)
    temp_arr = ����_������.items()(���� - 1)
    �����������_������� = Int((2 * Rnd) + 1)
    Select Case �����������_�������
        Case 1
            temp_arr(1) = "-1"
        Case 2
            temp_arr(1) = "1"
    End Select
    Key = temp_arr(6)
    ������� = CLng(Date)
    Range("P1") = (������� - Asc(Key) - 43000) * 10 + temp_arr(1)
End Sub

Function �������()
    temp = Right(Worksheets("����1").Range("P1"), 1)
    If temp = 9 Then
        side = -1
        �������� = 2
    Else
        If temp = 1 Then
            side = 1
            �������� = 4
        End If
    End If
    st1 = Range("P1") - side
    st2 = st1 / 10
    st3 = st2 + 43000
    st4 = CLng(Date)
    ������� = Chr(CLng(Date) - 43000 - (Range("P1") - side) / 10) & ��������
    Worksheets("����1").Range("P2") = �������
End Function

Sub ��������()
    '�������� �� �� ���� ����� ��������� ��� ����� �������
    
    Range("a1") = Range("a1") + 1
    If Range("a1") > 3 Then
        MsgBox "���������"
        Exit Sub
    End If
    
    Set ����_������ = CreateObject("Scripting.Dictionary")
    Set ����_������ = CreateCoinKit
    Call ���������_���_�����
    ����_� = ���������������("L", Range("m1"))
    ����_� = ���������������("R", Range("m1"))
    
    temp = �������() '��� � �����
    ����� = Left(temp, 1)
    ��� = Right(temp, 1)
    temp_arr = ����_������.Item(�����)
    ����� = temp_arr(1)

    For Each Key In ����_������
        temp_arr = ����_������.Item(Key) '������� ������� ����� �� �����
        If temp_arr(1) = "L" Then
            cnt_A = cnt_A + 1
        Else
            If temp_arr(1) = "R" Then
                cnt_B = cnt_B + 1
            End If
        End If
    Next Key
    
    If cnt_A = cnt_B Then
        flg_���������_��������� = 1
        If ����� = "L" Then
            flg = "A" & ���
            flg_���������_������� = 1
        Else
            If ����� = "R" Then
                If ��� = 2 Then
                    flg = "A" & ��� + 2
                    flg_���������_������� = 1
                End If
                If ��� = 4 Then
                    flg = "A" & ��� - 2
                    flg_���������_������� = 1
                End If
             Else
                If ����� = "T" Then
                    
                End If
             End If
        End If
        
    Else
        flg_��������_��������� = 0
        If cnt_A > cnt_B Then
            flg = "A4"
        Else
            If cnt_A < cnt_B Then
                flg = "A2"
            End If
        End If
    End If
    
    '���������� ������� ����� ����� � ����� ����� ��� �����������
    Select Case flg
        Case "A2"
            �����_���_�������� = "����_1"
            Worksheets("�����").Range("����_1�").Clear
            Worksheets("�����").Range("����_1�").Interior.Color = vbWhite
            Worksheets("�����").Range("����_1�").Clear
            Worksheets("�����").Range("����_1�").Interior.Color = vbWhite
            
            Worksheets("����1").Range(����_�).Copy _
                    Destination:=Worksheets("�����").Range("����_1�")
            Worksheets("����1").Range(����_�).Copy _
                    Destination:=Worksheets("�����").Range("����_1�")
        Case "A4"
            �����_���_�������� = "����_3"
            Worksheets("�����").Range("����_3�").Clear
            Worksheets("�����").Range("����_3�").Clear
            Worksheets("�����").Range("����_3�").Interior.Color = vbWhite
            Worksheets("�����").Range("����_3�").Interior.Color = vbWhite
                        
            Worksheets("����1").Range(����_�).Copy _
                    Destination:=Worksheets("�����").Range("����_3�")
            Worksheets("����1").Range(����_�).Copy _
                    Destination:=Worksheets("�����").Range("����_3�")
        Case Else
            �����_���_�������� = "����_2"
            Worksheets("�����").Range("����_2�").Clear
            Worksheets("�����").Range("����_2�").Clear
            Worksheets("�����").Range("����_2�").Interior.Color = vbWhite
            Worksheets("�����").Range("����_2�").Interior.Color = vbWhite
            
            Worksheets("����1").Range(����_�).Copy _
                    Destination:=Worksheets("�����").Range("����_2�")
            Worksheets("����1").Range(����_�).Copy _
                    Destination:=Worksheets("�����").Range("����_2�")
    End Select
    
    Worksheets("�����").Range(�����_���_��������).Copy Destination:=Worksheets("����1").Range("��")
    If flg_���������_��������� = 1 Then
        Call ���������_���������(flg_���������_�������)
    End If
    
    Range("b1") = Range("c1")
    Range("c1") = Range("d1")
    Range("d1") = Range("m1")
End Sub

Sub �������_����������()
    Set ����_������ = CreateObject("Scripting.Dictionary")
    Call CreateCoinKit
    Call ���������_���_�����
    adr = ���������������("R", Range("m1"))
    Call ����������(adr)
    adr = ���������������("L", Range("m1"))
    Call ����������(adr)
End Sub

Sub ����������(adr)
    Set Rng = Range(adr)
    n_cln = Rng.Columns.Count
    n_rw = Rng.Rows.Count
    
'    Dim temp_arr(1 To 6)
    Set ���_��� = CreateObject("Scripting.Dictionary")
    
    For q = n_cln To 1 Step -1
        For i = n_rw To 1 Step -1
'            rng.Cells(i, q).Select
            prosmotr = Rng.Cells(i, q)
        
            If Rng.Cells(i, q) = "" Then
                For R = i - 1 To 1 Step -1
                    If Rng.Cells(R, q) <> "" Then
                        Key = Rng.Cells(R, q).Value
                        Item = ����_������.Item(Key)
                        ����� = Item(1)
                        temp_arr = ����(�����, 1, q)
                        Item(2) = temp_arr(1)
                        If ���_���.exists(Key) = True Then
                            ���_���.Item(Key) = Item
                        Else
                            ���_���.Add Key, Item
                        End If
                        Call ������_�������_�����(���_���)
                        ���_���.RemoveAll
                    End If
                Next R
            End If
        
        
        Next i
    Next q
End Sub

Sub ���������_���������(flg_���������_�������)

    ���������_����� = Worksheets("����1").Range("m1")
    
    adr_L = ���������������("L", Range("m1"))
    adr_R = ���������������("R", Range("m1"))
    
    ����_��� = Worksheets("�����").Range("a30").Interior.Color
    ����_��� = Worksheets("�����").Range("a31").Interior.Color
        
    Select Case ���������_�����
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
            Call ��������_������(adr_L, adr_R, ����_���, ����_���)
            
            For Each cell In Range("����")
                If cell.Value <> "" Then
                    cell.Interior.Color = vbBlue
                End If
            Next cell
            
        Case -1
            Call ��������_������(adr_L, adr_R, ����_���, ����_���)
                                    
            For Each cell In Range("����")
                If cell.Value <> "" Then
                    cell.Interior.Color = vbBlue
                End If
            Next cell
            
        End Select
        
'        if
'    If ���������_����� = 0 Then
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
'            If ���������_����� = 1 Or ���������_����� = -1 Then
'                For Each cell In Range(adr_L)
'                    If cell.Value <> "" Then
'                        If cell.Interior.Color = vbYellow Then
'                            cell.Interior.Color = ����_L
'                        End If
'                    End If
'                Next cell
'                For Each cell In Range(adr_R)
'                    If cell.Value <> "" Then
'                        If cell.Interior.Color = vbYellow Then
'                            cell.Interior.Color = ����_R
'                        End If
'                    End If
'                Next cell
'
'                For Each cell In Range("����")
'                    If cell.Value <> "" Then
'                        cell.Interior.Color = vbBlue
'                    End If
'                Next cell
'            End If
'        End If
        
'    Call ��������_������(adr_L, adr_R, ����_L, ����_R)

    '����� ����� ����� ����� ��� �� ������ ����������
    If flg_���������_������� = 1 Then Worksheets("�����").Range("a29") = 1
    End Sub

Sub ��������_������(adr_L, adr_R, ����_L, ����_R)
'    ����������� = Range("a1")
        
    For Each cell In Range(adr_L)
        If cell.Value <> "" Then
            If cell.Interior.Color = vbYellow Then
                cell.Interior.Color = ����_L
            End If
        End If
    Next cell
    
    For Each cell In Range(adr_R)
        If cell.Value <> "" Then
            If cell.Interior.Color = vbYellow Then
                cell.Interior.Color = ����_R
            End If
        End If
    Next cell
    
End Sub

Sub �����������()
    Set Rng = Selection
    For Each cell In Rng
        If cell.Value <> "" Then cell.Interior.Color = vbBlue
    Next cell
End Sub

'Sub ����_����(�����_����, ������_����)
'    arr_����_����_��� = Array("����_1�", _
'        "����_1�", _
'        "����_2�", _
'        "����_2�", _
'        "����_3�", _
'        "����_3�")
'    For i = 0 To UBound(arr_����_����_���)
'        adr = Worksheets("�����").Range(arr_����_����_���(i)).Address
'        Set Rng = Range(adr)
'
'        Set Rng = Rng.Cells(Rng.Cells.Count).Offset(1, -2)
''        Set Rng =
''        Rng.Select
'        Set Rng = Rng.Resize(1, 3)
''        Rng.Select
'        If i Mod 2 = 0 Then
'            Rng.Interior.Color = �����_����
'        Else
'            Rng.Interior.Color = ������_����
'        End If
'    Next i
'
'End Sub
Sub the_bot()

    adr_L = ���������������("L", Range("m1"))
    adr_R = ���������������("R", Range("m1"))
    
    num_L = Range(adr_L).Count
    num_R = Range(adr_R).Count
    num_T = Range("����").Count
    
    Select Case Worksheets("����1").Range("levl")
        
        Case 3
            
        Case 2
        Case 1
        Case 0
        Case Else
    End Select
End Sub

Sub ��������_�����()
    a = Left(�������, 1)
    b = Range("q6")
    
    If a = b Then
        Range("����") = Range("����") + 1
    Else
        Range("����").Clear
        Range("n1", "n24").Clear
        Exit Sub
    End If
    
    For i = 24 To 1 Step -1
        If Range("N" & i) = "" Then
            adr = Range("N" & i).Address
            GoTo ����������_�����
        End If
    Next i
    
����������_�����:
    Range("q6").Cut Destination:=Worksheets("����1").Range(adr)
End Sub
