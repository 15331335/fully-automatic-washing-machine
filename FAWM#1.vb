Sub step1()
    ActiveSheet.Name = "全国"
    ActiveSheet.UsedRange.AutoFilter Field:=5, Criteria1:="=*广东*"
    Call step2
End Sub

Sub step2()
    ActiveSheet.UsedRange.Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    ActiveSheet.Name = "广东"
    Call step3
End Sub

Sub step3()
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
    Call step4
End Sub

Sub step4()
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
    Call step5
End Sub

Sub step5()
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
    Call step6
End Sub

Sub step6()
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
    Call step7
End Sub

Sub step7()
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
End Sub
