Attribute VB_Name = "SumPurcheseReq"
Option Explicit

Sub SumPuchaseRequest()

Worksheets("��z���ʌ���V�[�g").Activate

'�Z���[���̏��i�R�[�h���d���Ȃ��ŃR�s�[
'URL http://www.eurus.dti.ne.jp/yoneyama/Excel/vba/vba_jyufuku.html
Worksheets("�Z���[��").Range("C2:D" & Worksheets("�Z���[��").UsedRange.Rows.Count).AdvancedFilter xlFilterCopy, CriteriaRange:=Range("C2:C" & Worksheets("�Z���[��").UsedRange.Rows.Count), Unique:=True, CopyToRange:=Range("G2")

'��z�����̍��v�����ƁA��z�˗����̍��v������
Dim r As Range, CodeRange As Range, Endrow As Long
Endrow = Cells(1, 7).SpecialCells(xlLastCell).Row
Set CodeRange = Range(Cells(2, 7), Cells(Endrow, 7))

For Each r In CodeRange
    r.Offset(0, -1).Value = CountMallOrder(r.Value)
    r.Offset(0, 2).Value = SumRequestQuantity(r.Value, True)
Next

'�����̎�z�˗����ʂ̍��v���Z�o
'�����̏��i�R�[�h������
Worksheets("����").Range("C2:D" & Worksheets("����").UsedRange.Rows.Count).AdvancedFilter xlFilterCopy, CriteriaRange:=Range("C2:C" & Worksheets("�Z���[��").UsedRange.Rows.Count), Unique:=True, CopyToRange:=Range("G" & Endrow + 1)

Dim WholeSaleJanRange As Range
Set WholeSaleJanRange = Range(Cells(Endrow + 1, 7), Cells(2, 7).End(xlDown))

WholeSaleJanRange.Offset(0, -1).Value = "V"

For Each r In WholeSaleJanRange
    r.Offset(0, 2).Value = SumRequestQuantity(r.Value, False)
Next

End Sub

Private Function CountMallOrder(ByVal Code As String) As String

Dim Counter(3) As Long, Mall As String, FoundCode As String, Endrow As Long, k As Long, Mark As String

Endrow = Worksheets("�Z���[��").UsedRange.Rows.Count

For k = 2 To Endrow
    
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

CountMallOrder = Mark

End Function

Private Function SumRequestQuantity(ByVal Code As String, ByVal IsSeller As Boolean) As Long

    Dim TargetSheet As Worksheet, Endrow As Long, TmpSum As Long

    '���ʗ�͂ǂ����E��Ȃ̂ŁA�ΏۃV�[�g��؂芷���邾���ł悢�B
    Set TargetSheet = Worksheets(IIf(IsSeller, "�Z���[��", "����"))
    
    Endrow = TargetSheet.UsedRange.Rows.Count
     
    Dim TargetRange As Range, QuantityRange As Range
    Set TargetRange = TargetSheet.Range("C2").Resize(Endrow - 1, 1)
    Set QuantityRange = TargetSheet.Range("E2").Resize(Endrow - 1, 1)
    
    TmpSum = WorksheetFunction.SumIf(TargetRange, Code, QuantityRange)
    
    SumRequestQuantity = TmpSum

End Function



