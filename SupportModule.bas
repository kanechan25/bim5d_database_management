Attribute VB_Name = "SupportModule"

Sub Import1(RibbonControl As IRibbonControl)
UpdateData.Show
End Sub

Sub vlookupCate()
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(R2C17,R4C3:R2000C4,2,FALSE),"""")"
End Sub
Sub vlookupType()
    Range("T2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(R2C19,R4C5:R2000C6,2,FALSE),"""")"
End Sub
Sub vlookupSub()
    Range("V2").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(R2C21,R4C7:R2002C8,2,FALSE),"""")"
End Sub
Sub BB()
    ''''BB CHO AO (SHEET A VA S)
    Range("BB4").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-13]="""","""",LEFT(RC[-13],SEARCH(""."",RC[-13])-1))"
    Range("BB4").Select
    Selection.AutoFill Destination:=Range("BB4:BB1000"), Type:=xlFillDefault

    ''''BB CHO AO (SHEET MEPF)
    Range("BB4").Select
    ActiveCell.Formula2R1C1 = _
        "=IFS(RC[-11]="""","""",LEN(RC[-11])-LEN(SUBSTITUTE(RC[-11],""."",""""))=3,LEFT(RC[-11],vitri(""."",RC[-11],3)-1),LEN(RC[-11])-LEN(SUBSTITUTE(RC[-11],""."",""""))=2,LEFT(RC[-11],vitri(""."",RC[-11],2)-1),TRUE,LEFT(RC[-11],vitri(""."",RC[-11],1)-1))"
    Range("BB4").Select
    Selection.AutoFill Destination:=Range("BB4:BB2000"), Type:=xlFillDefault
End Sub
Sub Duplicate_ID_and_Name()
''''To mau do neu co ID va Family Name nao bi duplicate
    lrB = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
    ActiveSheet.Range("I4:I" & lrB).Select
    ActiveSheet.Range("K4:K" & lrB).Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
     ActiveWindow.ScrollColumn = 1
End Sub
''=vitri(kytu,chuoi,lan xuat hien thu n)
Function vitri(kytu As String, chuoi As String, num As Integer) As Variant
    Dim i As Integer
    Dim dem As Integer
    dem = 0
    For i = 1 To Len(chuoi)
     If Mid(chuoi, i, 1) = kytu Then
        dem = dem + 1
        If dem = num Then
        vitri = i
        End If
     End If
    Next i
End Function
Sub sort_system_family()
''''Tim dong cuoi cung cua system family
    Dim i As Long
    ActiveSheet.Range("BA1").Value = "=MATCH(""Loadable"",B:B,0)-1"
    i = ActiveSheet.Range("BA1").Value
''''Sort vung du lieu System Family
    ActiveSheet.Range("B4:K" & i).Select ''Sheet A va S la K, sheet MEPF la M
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add2 Key:=Range("C4:C" & i), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("B4:K" & i)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

