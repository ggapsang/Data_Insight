VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LookupInput_Form 
   Caption         =   "UserForm1"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6015
   OleObjectBlob   =   "LookupInput_Form.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "LookupInput_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lstLookupValues_Click()

End Sub

Private Sub txtHeaderRowSource_Change()

End Sub

Private Sub txtHeaderTargeSource_Change()

End Sub


Private Sub UserForm_Initialize()
    Dim tgtWb As Workbook
    Dim tgtWs As Worksheet
    Dim tgtHeaderRow As Long
    Dim i As Long
    Dim lastCol As Long

    ' Set the target workbook and worksheet
    Set tgtWb = ActiveWorkbook
    Set tgtWs = tgtWb.ActiveSheet
    tgtHeaderRow = InputBox("해더 행 번호 : ")

    lastCol = tgtWs.Cells(tgtHeaderRow, tgtWs.Columns.count).End(xlToLeft).Column

    For i = 1 To lastCol
        lstLookupValues.AddItem tgtWs.Cells(tgtHeaderRow, i).value
    Next i
End Sub

Private Sub cmdSubmit_Click()

    Dim selectedItems As Collection
    Dim item As Variant
    Dim i As Long

    Set selectedItems = New Collection

    For i = 0 To lstLookupValues.ListCount - 1
        If lstLookupValues.Selected(i) Then
            selectedItems.Add lstLookupValues.List(i)
        End If
    Next i

    If selectedItems.count = 0 Then
        MsgBox "찾을 값은 한개 이상 입력"
    ElseIf txtHeaderRowSource.value = "" Or txtHeaderRowTarget.value = "" Or txtKeyValue.value = "" Then
        MsgBox "해더 행 번호와 키 값 다시 확인"
    Else
        UpdateTargetWorksheet selectedItems, CLng(txtHeaderRowSource.value), CLng(txtHeaderRowTarget.value), txtKeyValue.value
        Unload Me
    End If

End Sub
