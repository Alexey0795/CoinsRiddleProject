Attribute VB_Name = "b_bot"
Sub the_bot_main()

    ����_��� = Worksheets("�����").Range("a30").Interior.Color
    ����_��� = Worksheets("�����").Range("a31").Interior.Color
    
    ��� = Range("a1")
    Select Case ���
        Case ""
            GoTo ������_���:
        Case 1
            GoTo ������_���:
        
        Case Else
            Stop
    End Select
    Range("e29").Select
    
������_���:
    '1-�� ���
    ������ = Range("����").Address
    Call �������_������2(������, 4, vbYellow)
    Call ������_����
    Call �������_������2(������, 4, vbYellow)
    Call ������_�����
    Call ��������
    
������_���:
    '2-�� ���
    If Range("m1") = 0 Then
        ������ = Range("����_��_�_���").Address
        Call �������_������2(������, 4, vbBlue)
        Call ������_����
        ������ = Range("����_��_�_���").Address
        Call �������_������2(������, 3, vbBlue)
        Call ������_����
        ������ = Range("����").Address
        Call �������_������2(������, 2, vbYellow)
        Call ������_����
        Call �������_������2(������, 1, vbYellow)
        Call ������_�����
    Else
        If Range("m1") = -1 Then
            ���_���� = "L"
            ���_���� = "R"
        Else
            ���_���� = "R"
            ���_���� = "L"
        End If
        
        ������ = ���������������(���_����, Range("m1"))
        Call �������_������2(������, 2, ����_���)
        If ���_���� = "R" Then
            Call ������_����
            Call �������_������2(Range("����").Address, 2, vbBlue)
            Call ������_�����
        Else
            Call ������_�����
            Call �������_������2(Range("����").Address, 2, vbBlue)
            Call ������_����
        End If
        Call �������_������2(������, 1, ����_���)
        Call ������_����
                
        ������ = ���������������(���_����, Range("m1"))
        Call �������_������2(������, 2, ����_���)
        Call ������_����
        Call �������_������2(������, 1, ����_���)
        If ���_���� = "L" Then
            Call ������_����
            Call �������_������2(Range("����").Address, 2, vbBlue)
            Call ������_�����
        Else
            Call ������_�����
            Call �������_������2(Range("����").Address, 2, vbBlue)
            Call ������_����
        End If
        Call �������_������2(������, 1, ����_���)
        Call ������_����
    End If
    
    Call ��������
    
    '3-�� ���
    '��������� �����
    '��������� �����
    Set ����_������ = CreateObject("Scripting.Dictionary")
    Set ����_������ = CreateCoinKit
    Call ���������_���_�����
    For Each Key In ����_������
        Item = ����_������.Item(Key)
        ���� = Item(3)
        Select Case ����
            Case vbBlue
                cnt_��� = cnt_��� + 1
            Case vbYellow
                cnt_��� = cnt_��� + 1
            Case ����_���
                cnt_��� = cnt_��� + 1
            Case ����_���
                cnt_��� = cnt_��� + 1
        End Select
    Next Key
    
    
'    adr = ���������������("L", Range("m1"))
'    Set Rng = Range(adr)
'    For Each cell In Rng
'        ���� = cell.Interior.Color
'        Select Case ����
'            Case vbBlue
'                cnt_��� = cnt_��� + 1
'            Case vbYellow
'                cnt_��� = cnt_��� + 1
'            Case ����_���
'                cnt_��� = cnt_��� + 1
'            Case ����_���
'                cnt_��� = cnt_��� + 1
'        End Select
'    Next cell
''____________________________________________
'    adr = ���������������("R", Range("m1"))
'    Set Rng = Range(adr)
'    For Each cell In Rng
'        ���� = cell.Interior.Color
'        Select Case ����
'            Case vbBlue
'                cnt_��� = cnt_��� + 1
'            Case vbYellow
'                cnt_��� = cnt_��� + 1
'            Case ����_���
'                cnt_��� = cnt_��� + 1
'            Case ����_���
'                cnt_��� = cnt_��� + 1
'        End Select
'    Next cell
''____________________________________________
'    Set Rng = Range("����")
'    For Each cell In Rng
'        ���� = cell.Interior.Color
'        Select Case ����
'            Case vbBlue
'                cnt_��� = cnt_��� + 1
'            Case vbYellow
'                cnt_��� = cnt_��� + 1
'            Case ����_���
'                cnt_��� = cnt_��� + 1
'            Case ����_���
'                cnt_��� = cnt_��� + 1
'        End Select
'    Next cell
    
    ����_1 = Range("c1")
    a = DoEvents
    ����_2 = Range("m1").Value
'==================================================
    If cnt_��� = 8 Then
    a = DoEvents
        If ����_1 = 1 And ����_2 = 1 Then
            ����� = Range("����_��_�_���").Address
            ������ = Range("����_��_�_���").Address
        Else
            ����� = Range("����_��_�_���").Address
            ������ = Range("����_��_�_���").Address
        End If
            
            ������ = ������
            Call �������_������2(������, 1, ����_���)
            ���_������� = Selection
            Call �������_������2(������, 1, vbBlue)
            Call ������_����
            Call �������_������2(������, 1, ����_���)
            Call �����������
            Call ������_����
            
            ������ = �����
            Call �������_������2(������, 1, ����_���)
            ���_������� = Selection
            Call �������_������2(������, 2, vbBlue)
            Call ������_����
            Call �������_������2(������, 1, ����_���)
            Call �����������
            Call ������_����
            
            Call ��������
            a = DoEvents
            
            If Range("m1") = 0 Then
                Range("����").Find(���_�������).Select
            Else
                Range("����_��_�_���").Find(���_�������).Select
            End If
            Call ������_�����
            Exit Sub
'        Else
'            Stop
'        End If
    Else
        If ����_1 = 0 And ����_2 = 0 Then
            Call �������_������2(Range("����").Address, 1, vbYellow)
            Call ������_�����
            Exit Sub
        End If
        
        If ����_1 = 0 And ����_2 <> 0 Then
            If cnt_��� > cnt_��� Then
                
            End If
            Call �������_������2(Range("����").Address, 1, vbYellow)
            Call ������_�����
            Exit Sub
        End If
        Stop
    End If
'==================================================
    Stop
End Sub

Sub �������_������2(������, �������, Optional ������_������ = "")
    '�������� ������ �� �����
    Dim arr_��������()
    Set ���_�������� = CreateObject("Scripting.Dictionary")
    
    Set Rng = Range(������)
    
    For Each cell In Rng
        
'        arr_item = ����_������.Item(Key)
            
            If ������_������ <> "" Then
                If cell.Interior.Color = ������_������ Then
                    flg_�������_����� = 1
                Else
                    flg_�������_����� = 0
                End If
            Else
                flg_�������_����� = 1
            End If
            
            If flg_�������_����� = 1 Then
                cnt_������� = cnt_������� + 1
                
                ReDim Preserve arr_��������(1 To cnt_�������)
                arr_��������(cnt_�������) = cell.Address
                
                If cnt_������� >= ������� Then
'                    Set �������_������ = ���_��������
                    GoTo ���������_�����
                End If
            End If
    
    Next cell
���������_�����:
    For i = 1 To UBound(arr_��������)
        ������ = ������ & arr_��������(i) & ","
    Next i
    ������ = Left(������, Len(������) - 1)
    
    Worksheets("����1").Range(������).Select
'    Set Rng = Range(������)
'    Set ���_�������� = CreateObject("Scripting.Dictionary")
'    For Each Cell In Rng
'        While n < �������
'            If Cell.Interior.Color = ������_������ Then
'                Key = Cell.Value
'                ���_��������.Add Key, 1
'            End If
'        Loop
'    Next Cell
'
'    Set �������_������ = ���_��������
End Sub

