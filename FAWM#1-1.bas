Attribute VB_Name = "模块1"
Sub 全自动洗衣机()
    ActiveSheet.Name = "全国"
    ActiveSheet.UsedRange.AutoFilter Field:=5, Criteria1:="=*广东*"
    
    ActiveSheet.UsedRange.Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    ActiveSheet.Name = "广东"
    
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
    
    ActiveSheet.Name = "统计"
    Columns("A:A").EntireColumn.AutoFit
    Range("D1") = "地市"
    Range("E1") = "数量"
    
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-3],RC[-1])"
    ir = Range("d65536").End(xlUp).Row
    Selection.AutoFill Destination:=Range("E2:E" & ir)
    
    Range("D" & ir + 1) = "合计"
    Range("E" & ir + 1).Select
    ActiveCell.Formula = "=SUM(E2:E" & ir & ")"
    
    Sheets("广东").Select
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "标记"
    
    Sheets("广东").Select
    Rows("1:1").Copy Sheets("标记").Range("A1")
    i = 2
    For r = 2 To ActiveSheet.UsedRange.Rows.Count
        For c = 1 To ActiveSheet.UsedRange.Columns.Count
            If Cells(r, c).Interior.ColorIndex = 6 Then
                Cells(r, c).EntireRow.Copy Sheets("标记").Range("A" & i)
                i = i + 1
                Exit For
            End If
        Next
    Next
    
    Sheets("标记").Select
    Columns("E:E").Select
    Selection.Copy
    Sheets("统计").Select
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
    
    Range("J1") = "地市"
    Range("K1") = "数量"
    
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-3],RC[-1])"
    Range("K2").Select
    ir = Range("j65536").End(xlUp).Row
    Selection.AutoFill Destination:=Range("K2:K" & ir)
    Range("J" & ir + 1) = "合计"
    Range("K" & ir + 1).Select
    ActiveCell.Formula = "=SUM(K2:K" & ir & ")"
    
    ir = Range("j65536").End(xlUp).Row
    Range("J2:K" & ir - 1).Select
    Range("K" & ir - 1).Activate
    ActiveWorkbook.Worksheets("统计").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("统计").Sort.SortFields.Add2 Key:=Range("K" & ir - 1), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("统计").Sort
        .SetRange Range("J2:K" & ir - 1)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("M1").Select
    ActiveCell.FormulaR1C1 = "不合格率"
    Range("N1").Select
    Application.CutCopyMode = False
    ir1 = Range("e65536").End(xlUp).Row
    ir2 = Range("k65536").End(xlUp).Row
    ActiveCell.Formula = "=K" & ir2 & "/E" & ir1
    Range("N1").Select
    Selection.NumberFormatLocal = "0.00%"
    
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "案件类别选择错误"
    Range("M3").Select
    ActiveCell.FormulaR1C1 = "账号录入、涉案电话不规范"
    Range("M4").Select
    ActiveCell.FormulaR1C1 = "简要案情过于简单"
    
    Sheets("标记").Select
    c = 0
    For r = 2 To ActiveSheet.UsedRange.Rows.Count
        If Cells(r, 4).Interior.ColorIndex = 6 Then
            c = c + 1
        End If
    Next
    Sheets("统计").Range("N2") = c
    
    c = 0
    For r = 2 To ActiveSheet.UsedRange.Rows.Count
        If Cells(r, 18).Interior.ColorIndex = 6 Then
            c = c + 1
        End If
    Next
    Sheets("统计").Range("N4") = c
    
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
    Sheets("统计").Range("N3") = c
    Sheets("统计").Select
    Columns("M:M").EntireColumn.AutoFit
End Sub

