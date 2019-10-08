Attribute VB_Name = "ģ��1"
Sub ȫ�Զ�ϴ�»�()
    ActiveSheet.Name = "ȫ��"
    ActiveSheet.UsedRange.AutoFilter Field:=5, Criteria1:="=*�㶫*"
    
    ActiveSheet.UsedRange.Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    ActiveSheet.Name = "�㶫"
    
    Columns("E:E").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    
    Range("B2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=MID(RC[-1],4,2)"
    Range("B2").Select
    ir = Range("a65536").End(xlUp).Row
    Selection.AutoFill Destination:=Range("B2:B" & ir)
    
    Columns("B:B").Select
    Selection.Copy
    Range("D1").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Columns("D:D").Select
    ActiveSheet.Range("D1:D" & ir).RemoveDuplicates Columns:=1, Header:=xlNo
    
    ActiveSheet.Name = "ͳ��"
    Columns("A:A").EntireColumn.AutoFit
    Range("D1") = "����"
    Range("E1") = "����"
    
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-3],RC[-1])"
    ir = Range("d65536").End(xlUp).Row
    Selection.AutoFill Destination:=Range("E2:E" & ir)
    
    Range("D" & ir + 1) = "�ϼ�"
    Range("E" & ir + 1).Select
    ActiveCell.Formula = "=SUM(E2:E" & ir & ")"
    
    Sheets("�㶫").Select
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "���"
    
    Sheets("�㶫").Select
    Rows("1:1").Copy Sheets("���").Range("A1")
    i = 2
    For r = 2 To ActiveSheet.UsedRange.Rows.Count
        For c = 1 To ActiveSheet.UsedRange.Columns.Count
            If Cells(r, c).Interior.ColorIndex = 6 Then
                Cells(r, c).EntireRow.Copy Sheets("���").Range("A" & i)
                i = i + 1
                Exit For
            End If
        Next
    Next
    
    Sheets("���").Select
    Columns("E:E").Select
    Selection.Copy
    Sheets("ͳ��").Select
    Range("G1").Select
    ActiveSheet.Paste
    Columns("G:G").EntireColumn.AutoFit
    Range("H2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=MID(RC[-1],4,2)"
    Range("H2").Select
    ir = Range("g65536").End(xlUp).Row
    Selection.AutoFill Destination:=Range("H2:H" & ir)
    
    Columns("H:H").Select
    Selection.Copy
    Range("J1").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Columns("J:J").Select
    Application.CutCopyMode = False
    ActiveSheet.Range("$J$1:$J$341").RemoveDuplicates Columns:=1, Header:=xlNo
    
    Range("J1") = "����"
    Range("K1") = "����"
    
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-3],RC[-1])"
    Range("K2").Select
    ir = Range("j65536").End(xlUp).Row
    Selection.AutoFill Destination:=Range("K2:K" & ir)
    Range("J" & ir + 1) = "�ϼ�"
    Range("K" & ir + 1).Select
    ActiveCell.Formula = "=SUM(K2:K" & ir & ")"
    
    ir = Range("j65536").End(xlUp).Row
    Range("J2:K" & ir - 1).Select
    Range("K" & ir - 1).Activate
    ActiveWorkbook.Worksheets("ͳ��").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ͳ��").Sort.SortFields.Add2 Key:=Range("K" & ir - 1), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ͳ��").Sort
        .SetRange Range("J2:K" & ir - 1)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "���ϸ���"
    Range("N1").Select
    Application.CutCopyMode = False
    ir1 = Range("e65536").End(xlUp).Row
    ir2 = Range("k65536").End(xlUp).Row
    ActiveCell.Formula = "=K" & ir2 & "/E" & ir1
    Range("N1").Select
    Selection.NumberFormatLocal = "0.00%"
    
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "�������ѡ�����"
    Range("M3").Select
    ActiveCell.FormulaR1C1 = "�˺�¼�롢�永�绰���淶"
    Range("M4").Select
    ActiveCell.FormulaR1C1 = "��Ҫ������ڼ�"
    
    Sheets("���").Select
    c = 0
    For r = 2 To ActiveSheet.UsedRange.Rows.Count
        If Cells(r, 4).Interior.ColorIndex = 6 Then
            c = c + 1
        End If
    Next
    Sheets("ͳ��").Range("N2") = c
    
    c = 0
    For r = 2 To ActiveSheet.UsedRange.Rows.Count
        If Cells(r, 18).Interior.ColorIndex = 6 Then
            c = c + 1
        End If
    Next
    Sheets("ͳ��").Range("N4") = c
    
    '20-26
    c = 0
    For r = 2 To ActiveSheet.UsedRange.Rows.Count
        If Cells(r, 20).Interior.ColorIndex = 6 Or _
            Cells(r, 21).Interior.ColorIndex = 6 Or _
            Cells(r, 22).Interior.ColorIndex = 6 Or _
            Cells(r, 23).Interior.ColorIndex = 6 Or _
            Cells(r, 24).Interior.ColorIndex = 6 Or _
            Cells(r, 25).Interior.ColorIndex = 6 Or _
            Cells(r, 26).Interior.ColorIndex = 6 Then
            c = c + 1
        End If
    Next
    Sheets("ͳ��").Range("N3") = c
    Sheets("ͳ��").Select
    Columns("M:M").EntireColumn.AutoFit
End Sub

