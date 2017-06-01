Attribute VB_Name = "SumProduct"
Option Explicit

Sub SumPuchaseRequest()

'�Z���[���̏��i�R�[�h���d���Ȃ��ŃR�s�[
'URL http://www.eurus.dti.ne.jp/yoneyama/Excel/vba/vba_jyufuku.html

Worksheets("�Z���[��").Range("C2:C48").AdvancedFilter xlFilterCopy, Unique:=True, CopyToRange:=Range("B2")

'��z�����𗪍��œ����
Dim r As Range, CodeRange As Range, EndRow As Long
EndRow = Cells(1, 2).SpecialCells(xlLastCell).Row
Set CodeRange = Range(Cells(2, 2), Cells(EndRow, 2))

'���i���̎�z�˗����ʂ̍��v���Z�o

For Each r In CodeRange
    r.Offset(0, -1).Value = BuildMark(r.Value)
Next

'�����̏��i�R�[�h������
Worksheets("����").Range("C2:C5").AdvancedFilter xlFilterCopy, Unique:=True, CopyToRange:=Range("B46")

Range("A2").End(xlDown).Offset(1, 1).Select
Range(Selection, Selection.End(xlDown)).Offset(0, -1).Select
Selection.Value = "V"
Range("A1").Select


'�����̎�z�˗����ʂ̍��v���Z�o


End Sub

Private Function BuildMark(ByVal Code As String) As String

Dim Counter(3) As Long, Mall As String, FoundCode As String, EndRow As Long, k As Long, Mark As String

EndRow = Worksheets("�Z���[��").UsedRange.Rows.Count

For k = 2 To EndRow
    
    FoundCode = Worksheets("�Z���[��").Cells(k, 3).Value
    
    If FoundCode = Code Then
    
        Mall = Worksheets("�Z���[��").Cells(k, 1).Value
        
        Select Case Mall
            Case "A"
                Counter(0) = Counter(0) + 1
            Case "R"
                Counter(1) = Counter(1) + 1
            Case "Y"
                Counter(2) = Counter(2) + 1
            Case Else
                Counter(3) = Counter(3) + 1
        End Select
    
    End If

Next

Mark = IIf(Counter(0) > 0, "A" & Counter(0), "")
Mark = Mark & IIf(Counter(1) > 0, "R" & Counter(1), "")
Mark = Mark & IIf(Counter(2) > 0, "Y" & Counter(2), "")
Mark = Mark & IIf(Counter(3) > 0, "SP" & Counter(3), "")

'���̂܂܂���A1�Ƃ����悤�ɁA1���̎����������o��̂ŁA1�͍폜
Mark = Replace(Mark, "1", "")

BuildMark = Mark

End Function
