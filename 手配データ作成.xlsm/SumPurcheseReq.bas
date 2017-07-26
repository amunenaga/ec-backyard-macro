Attribute VB_Name = "SumPurcheseReq"
Option Explicit

Sub SumPuchaseRequest()
'���i�ʂɎ�z�˗������W�v

Worksheets("��z���ʓ��̓V�[�g").Activate

'�Z���[���̏��i�R�[�h���R�s�[���āA�d���폜
If Worksheets("�Z���[��").Range("A2").Value = "" Then GoTo Vendor
Worksheets("�Z���[��").Range("C2:D" & Worksheets("�Z���[��").UsedRange.Rows.Count).Copy Destination:=Range("G2")
Range("A1").CurrentRegion.RemoveDuplicates Columns:=7, Header:=xlYes


'��z�����̍��v�����ƁA��z�˗����̍��v������
Dim r As Range, CodeRange As Range

Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange
    r.Offset(0, -1).Value = CountMallOrder(r.Value)
    r.Offset(0, 2).Value = SumRequestQuantity(r.Value, True)
Next

Vendor:

'�����̎�z�˗����ʂ̍��v���Z�o
'�����̏��i�R�[�h������
If Worksheets("����").Range("A2").Value = "" Then GoTo Quit
Worksheets("����").Range("C2:D" & Worksheets("����").UsedRange.Rows.Count).Copy Destination:=Range("G2").End(xlDown).Offset(1, 0)

'�R�s�[���ďd���폜
Range(Cells(2, 6).End(xlDown).Offset(1, 1), Cells(2, 7).End(xlDown).Offset(0, 1)).RemoveDuplicates 1, xlNo

Dim WholeSaleJanRange As Range
Set WholeSaleJanRange = Range(Cells(2, 6).End(xlDown).Offset(1, 1), Cells(2, 7).End(xlDown))
WholeSaleJanRange.Offset(0, -1).Value = "V"

For Each r In WholeSaleJanRange
    r.Offset(0, 2).Value = SumRequestQuantity(r.Value, False)
Next

Quit:

End Sub

Private Function CountMallOrder(ByVal Code As String) As String
'���[���ʂ̌������������BA2R�̓A�}�]��2���A�y�V1���̎�z�˗�����B

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

Mark = IIf(Counter(0) > 0, "A" & IIf(Counter(0) > 1, Counter(0), ""), "")
Mark = Mark & IIf(Counter(1) > 0, "R" & IIf(Counter(1) > 1, Counter(1), ""), "")
Mark = Mark & IIf(Counter(2) > 0, "Y" & IIf(Counter(2) > 1, Counter(2), ""), "")
Mark = Mark & IIf(Counter(3) > 0, "SP" & IIf(Counter(3) > 1, Counter(3), ""), "")

CountMallOrder = Mark

End Function

Private Function SumRequestQuantity(ByVal Code As String, ByVal IsSeller As Boolean) As Long
    '���i�R�[�h�ʂ̎�z�˗����ʂ����v����B
    
    Dim TargetSheet As Worksheet, EndRow As Long, TmpSum As Long

    '���ʗ�͂ǂ����E��Ȃ̂ŁA�ΏۃV�[�g��؂芷���邾���ł悢�B
    Set TargetSheet = Worksheets(IIf(IsSeller, "�Z���[��", "����"))
    
    EndRow = TargetSheet.UsedRange.Rows.Count
     
    Dim TargetRange As Range, QuantityRange As Range
    Set TargetRange = TargetSheet.Range("C2").Resize(EndRow - 1, 1)
    Set QuantityRange = TargetSheet.Range("E2").Resize(EndRow - 1, 1)
    
    TmpSum = WorksheetFunction.SumIf(TargetRange, Code, QuantityRange)
    
    SumRequestQuantity = TmpSum

End Function
