Attribute VB_Name = "Merge_EasyLookup_m"
Function GetWorkbook(ByVal sFullName As String) As Workbook
    Dim sFile As String
    Dim wbReturn As Workbook

    sFile = Dir(sFullName)

    On Error Resume Next
    Set wbReturn = Workbooks(sFile)

    If wbReturn Is Nothing Then
        Set GetWorkbook = Nothing
    Else
        Set GetWorkbook = wbReturn
    End If

    On Error GoTo 0
End Function

Function MatchKeyValue(keyValue As Variant, headerRow As Range) As Long
Dim cell As Range
    Dim i As Long
    
    For Each cell In headerRow.Cells
        If IsNumeric(keyValue) And IsNumeric(cell.value) Then
            If CDbl(keyValue) = CDbl(cell.value) Then
                MatchKeyValue = cell.Column
                Exit Function
            End If
        Else
            If CStr(keyValue) = CStr(cell.value) Then
                MatchKeyValue = cell.Column
                Exit Function
            End If
        End If
    Next cell

    MatchKeyValue = 0
End Function

Function MatchValue(lookupValue As Variant, headerRow As Range) As Long
    Dim cell As Range

    For Each cell In headerRow.Cells
        If IsNumeric(lookupValue) And IsNumeric(cell.value) Then
            If CDbl(lookupValue) = CDbl(cell.value) Then
                MatchValue = cell.Column
                Exit Function
            End If
        Else
            If CStr(lookupValue) = CStr(cell.value) Then
                MatchValue = cell.Column
                Exit Function
            End If
        End If
    Next cell

    MatchValue = 0
End Function

Sub Merge_EasyLookup()
    ' ���������� ���� ������ �� �ְ� ������
    LookupInputForm.Show
End Sub


Sub UpdateTargetWorksheet(selectedItems As Collection, srcHeaderRow As Long, tgtHeaderRow As Long, keyValueText As String)
    
    ' ��ũ�� �۾��� ��ũ�� Ȱ��ȭ ����
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim tgtWb As Workbook
    Dim tgtWs As Worksheet
    Dim tgtFile As Variant
    Dim srcKeyValueCol As Long
    Dim srcLookupValueCol() As Long
    Dim tgtKeyValueCol As Long
    Dim tgtLookupValueCol() As Long
    Dim srcKeyCell As Range
    Dim i As Long
    Dim j As Long
    Dim lastRow As Long
    Dim startTime As Double
    Dim finishTime As Double
    Dim tgtKeyCells As Object
    Dim rowNumber As Variant

     ' source worksheet ����
    Set srcWb = ActiveWorkbook
    Set srcWs = ActiveSheet

     ' target workbook ���� ����â ����
    tgtFile = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*), *.xls*", Title:="��� ���� ����", MultiSelect:=False)
    If tgtFile = False Then
        MsgBox "�ҷ��� ������ ���õ��� �ʾƼ� ��ũ�θ� �����մϴ�"
        Exit Sub
    End If

    ' target workbook�� �̹� �������� ���
    Set tgtWb = GetWorkbook(tgtFile)

    ' target workbook�� ���� ���� ���� ��� �ش� ������ ����
    If tgtWb Is Nothing Then
        Set tgtWb = Workbooks.Open(tgtFile)
    End If

    ' target worksheet ����
    Set tgtWs = tgtWb.ActiveSheet

    startTime = Timer

    ' source worksheet�� target worksheet�� Ű ���� �ش� �� ��� �ִ��� Ȯ��
    srcKeyValueCol = MatchKeyValue(keyValueText, srcWs.Rows(srcHeaderRow))
    tgtKeyValueCol = MatchKeyValue(keyValueText, tgtWs.Rows(srcHeaderRow))
    ReDim srcLookupValueCol(selectedItems.count - 1)
    ReDim tgtLookupValueCol(selectedItems.count - 1)

    For j = 1 To selectedItems.count
        srcLookupValueCol(j - 1) = MatchValue(selectedItems.item(j), srcWs.Rows(srcHeaderRow))
        tgtLookupValueCol(j - 1) = MatchValue(selectedItems.item(j), tgtWs.Rows(tgtHeaderRow))
    Next j

    If IsError(srcKeyValueCol) Or IsError(tgtKeyValueCol) Then
        MsgBox "�ش� �࿡�� Ű ���� ã�� ���� ��ũ�θ� �����մϴ�."
        Exit Sub
    End If


    ' targetworksheet�� ������ �� ã��(���� - sr no�� �� ä������ �ʰ� �߰��� ����� ������ ���� �� ����)
    lastRow = tgtWs.Cells(tgtWs.Rows.count, tgtKeyValueCol).End(xlUp).Row

    ' ��ųʸ� ����
    Set tgtKeyCells = New Collection
    On Error Resume Next
    For i = tgtHeaderRow + 1 To lastRow
        tgtKeyCells.Add i, CStr(tgtWs.Cells(i, tgtKeyValueCol).value)
    Next i
    On Error GoTo 0

    ' loop
    For Each srcKeyCell In srcWs.Range(srcWs.Cells(srcHeaderRow + 1, srcKeyValueCol), srcWs.Cells(srcWs.Rows.count, srcKeyValueCol).End(xlUp))
        ' ���Ͱ��� ���̰� �ݿ���
        If srcWs.Rows(srcKeyCell.Row).Hidden = False Then
            ' Ű ���� �������� target worksheet�� ������ ��ġ�ϴ����� ����
            On Error Resume Next
            rowNumber = tgtKeyCells(CStr(srcKeyCell.value))
            On Error GoTo 0
            If Not IsEmpty(rowNumber) Then
                 ' ���Ͱ��� ���̰� �ݿ���
                If tgtWs.Rows(rowNumber).Hidden = False Then
                    ' source worksheets�� target worksheets�� ������ ���Ͽ� ������
                    For j = LBound(srcLookupValueCol) To UBound(srcLookupValueCol)
                        If tgtWs.Cells(rowNumber, tgtLookupValueCol(j)).value <> srcWs.Cells(srcKeyCell.Row, srcLookupValueCol(j)).value Then
                            tgtWs.Cells(rowNumber, tgtLookupValueCol(j)).value = srcWs.Cells(srcKeyCell.Row, srcLookupValueCol(j)).value
                            tgtWs.Cells(rowNumber, tgtLookupValueCol(j)).Interior.Color = RGB(255, 165, 0) 'RGB ����'
                        End If
                    Next j
                End If
            End If
        End If
    Next srcKeyCell

    ' ��ũ�� �۾� �� ��ũ�� Ȱ��ȭ ����
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True

     ' �۾��ð� ����
    finishTime = Timer - startTime

    ' �ٲ� ������ ������ �ڵ����� ������ (����� ���� ��� ���� ���� ������ �������� �ʰų� ���� �ڵ����� �����Ϸ��� �Ʒ� �� �ڵ忡 ' ǥ�ø� ���� ��
    'tgtWb.Save
    MsgBox "�Ϸ�" & Format(Int(finishTime / 60), "0") & " min " & Format(finishTime Mod 60, "0.00") & " sec"

End Sub
