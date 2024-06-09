Sub 속성그룹코드_내림차순으로_필터링()
'
' 속성그룹코드_내림차순으로_필터링 매크로
'

'
    Rows("1:1").Select
    Range("V1").Activate
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("V1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub 항목찾은_후_줄띄우기_및_해더다시붙여넣기()
'
' 항목찾은_후_줄띄우기_및_해더다시붙여넣기 매크로
'

'
    Columns("V:V").Select
    Selection.Find(What:="02", After:=ActiveCell, LookIn:=xlFormulas2, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, MatchByte:=False, SearchFormat:=False).Activate
    Rows("4272:4272").Select
    Range("V4272").Activate
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("V4271").Select
    Selection.End(xlUp).Select
    Rows("1:1").Select
    Range("V1").Activate
    Selection.Copy
    Range("V2").Select
    Selection.End(xlDown).Select
    Range("V4273").Select
    Selection.End(xlToLeft).Select
    ActiveSheet.Paste
End Sub
Sub 열하나파기()
'
' 열하나파기 매크로
'

'
    Columns("Y:Y").Select
    Range("Y4273").Activate
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("Y4274").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC24, R1C21:R4271C21, 1, 0)"
    Range("Y4274").Select
    Selection.Copy
    Range("X4274").Select
    Selection.End(xlDown).Select
    Range("Y5127").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("Y4275:Y5127").Select
    Range("Y5127").Activate
    ActiveSheet.Paste
    Range("Y5126").Select
    Selection.End(xlUp).Select
    Range("Y4274").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("Y4273").Select
    Application.CutCopyMode = False
    Rows("4273:4273").Select
    Range("X4273").Activate
    Selection.AutoFilter
    Rows("4273:4273").Select
    Range("X4273").Activate
    Selection.AutoFilter
    Range("Y4273").Select
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("Y4273"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Columns("Y:Y").Select
    Range("Y4274").Activate
    Selection.Find(What:="#", After:=ActiveCell, LookIn:=xlFormulas2, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, MatchByte:=False, SearchFormat:=False).Activate
    Rows("4736:4736").Select
    Range("Y4736").Activate
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("4736:4738").Select
    Range("Y4736").Activate
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Range("Y4735").Select
    Selection.End(xlUp).Select
    Rows("4272:4273").Select
    Range("X4272").Activate
    Selection.Delete Shift:=xlUp
    Columns("Y:Y").Select
    Range("Y4272").Activate
    Selection.Delete Shift:=xlToLeft
    Range("X4272").Select
    Selection.End(xlUp).Select
    Rows("1:1").Select
    Range("W1").Activate
    Selection.AutoFilter
    Range("V1").Select
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("V1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("U1").Select
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("U1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
