Attribute VB_Name = "Append"
Sub AppendCode(ByVal Code As String, ByVal RangeName As String, Optional RowNumber As Variant)
'���X�g�ɃR�[�h��������

'���Ƀ��X�g�A�b�v�ς݂̃R�[�h�łȂ����`�F�b�N
If WorksheetFunction.CountIf(ThisWorkbook.Names(RangeName).RefersToRange, Code) > 0 Then Exit Sub

Dim N As Name
Set N = ThisWorkbook.Names(RangeName)

Dim CutLength As Integer
CutLength = InStr(2, N.Value, "!") - 2

Dim SheetName As String
SheetName = Mid(N, 2, CutLength)

If IsMissing(RowNumber) Then 'IsMissing�֐����g���A���肵����������Variant�łȂ��Ɣ���r�b�g���܂܂�Ȃ�

    'Yahoo6digits�V�[�g�̊Y���s����肷��
    On Error Resume Next
        
        RowNumber = WorksheetFunction.Match(CDbl(Code), yahoo6digit.Range("YahooCodeRange"), 0)
        If Err Then Exit Sub
    
    On Error GoTo 0
End If

'���i���R�[�h���R�s�[
With ThisWorkbook.Worksheets(SheetName)
    
    yahoo6digit.Rows(RowNumber).Copy Destination:=.Rows(.UsedRange.Rows.Count + 1)
    
    '���t�[�f�[�^�̕��̓O���[�œh��
    yahoo6digit.Range("A" & RowNumber & ":I" & RowNumber).Interior.ColorIndex = 15

End With

End Sub
