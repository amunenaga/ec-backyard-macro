Attribute VB_Name = "Append"
Sub AppendCode(Code As String, RangeName As String)
'���X�g�ɃR�[�h��������

'���Ƀ��X�g�A�b�v�ς݂̃R�[�h�łȂ����`�F�b�N
If WorksheetFunction.CountIf(ThisWorkbook.Names(RangeName).RefersToRange, Code) > 0 Then Exit Sub

Dim N As Name
Set N = ThisWorkbook.Names(RangeName)

Dim SheetName As String
SheetName = N.Value

Dim CutLength As Integer
CutLength = InStr(2, N.Value, "!") - 2

SheetName = Mid(SheetName, 2, CutLength)

Dim FindRow As Long
'���X�g�A�b�v����Ă��Ȃ���΁Ayahoo6digit����R�s�[
On Error Resume Next
    
    FindRow = WorksheetFunction.Match(CDbl(Code), yahoo6digit.Range("YahooCodeRange"), 0)
    If Err Then Exit Sub

On Error GoTo 0

With ThisWorkbook.Worksheets(SheetName)
    
    yahoo6digit.Rows(FindRow).Copy Destination:=.Rows(.UsedRange.Rows.Count + 1)
    yahoo6digit.Rows(FindRow).Interior.ColorIndex = 15

End With

End Sub
