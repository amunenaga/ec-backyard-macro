Attribute VB_Name = "BuildDataForMagic"
Option Explicit
Const OPERATOR_CODE As Integer = 329

Type Purchase

    Code As String
    ProductName As String
    
    VendorCode As Long
    VendorName As String
    
    UnitCost As Long
    
    PurchaseQuantity As Long
    RequireQuantity As Long
    
    WarehouseNumber As Integer
    
    IsPickup As Integer
    
    IsHold As Boolean
    HoldReason As String

End Type

Sub BuildPurcahseData()
    
Worksheets("��z���ʌ���V�[�g").Activate

Dim i As Long
For i = 2 To Range("A1").End(xlDown).Row

    Dim CurrentPurchase As Purchase
    CurrentPurchase = ReadPurchase(i)
    
    If Not CurrentPurchase.IsHold Then
        'Call WriteHoldList(CurrentPurchase)
    
        If CurrentPurchase.Code Like "######" Then
            Call WriteMagicTxt(CurrentPurchase)
        Else
            Call WriteMagicManualInput(CurrentPurchase)
        End If
        
    End If

Next

Worksheets("Magic�ꊇ�o�^").Columns("A:E").AutoFit
Worksheets("Magic����͗p").Columns("A:I").AutoFit

'Magic�ꊇ�o�^�V�[�g��V�K�u�b�N�ɃR�s�[�ACSV�ŕۑ�
Worksheets("Magic�ꊇ�o�^").Copy
ActiveSheet.Rows(1).Delete

Dim FileName As String
FileName = "\Magic�o�^�p" & Format(Date, "MMdd") & ".txt"

If Dir(ThisWorkbook.path & FileName) <> "" Then
    FileName = Replace(FileName, Format(Date, "MMdd"), Format(Date, "MMdd") & "-" & Format(Time, "hhmm"))
End If

ActiveWorkbook.SaveAs FileName:=ThisWorkbook.path & FileName, FileFormat:=xlText
ActiveWorkbook.Close

End Sub

Private Function ReadPurchase(ByVal Row As Long) As Purchase

'��z���ʌ���V�[�g����1�s��1�ϐ��ɓǂݍ���
Dim TmpPur As Purchase

With TmpPur
    .Code = Cells(Row, 7).Value  '�������̏��i�R�[�h�AJAN��6�P�^
    .ProductName = Cells(Row, 8).Value '���i���AJAN��z���̂ݕK�{
    
    .VendorCode = Cells(Row, 4).Value '��z��R�[�h
    .VendorName = Cells(Row, 5).Value '��z�於��
     
    .WarehouseNumber = IIf(Cells(Row, 6).Value = "V", "187", "180")  '�q�ɔԍ�

    .RequireQuantity = Cells(Row, 9).Value '��z�˗�����

    '�����ۗ��ɊY�����邩�`�F�b�N���āA�t���O�𗧂Ă�

    '�ۗ��Ƃ��Đ��ʂ𒍈ӏ����ŏ㏑�����Ă��邩�H
    If IsNumeric(Cells(Row, 1).Value) Then
        TmpPur.PurchaseQuantity = Cells(Row, 1).Value
    Else
        TmpPur.HoldReason = Cells(Row, 1).Value
        TmpPur.IsHold = True
    End If
    
    .UnitCost = Cells(Row, 10).Value '�����f�[�^�̗L��
    If .UnitCost = 0 Then
        TmpPur.IsHold = True
        TmpPur.HoldReason = "�����s��"
    End If
    
    If .VendorCode = 0 And .VendorName = "" Then '�d���悪�������Ă��邩
        TmpPur.IsHold = True
        TmpPur.HoldReason = "�d����s��"
    End If
    
    .IsPickup = GetPickupFlag(.VendorCode) '����Ŏ�z���邩
        
End With

ReadPurchase = TmpPur

End Function

Private Sub WriteMagicTxt(ByRef Purchase As Purchase)
    
    Dim WriteRow As Long, TargetSheet As Worksheet, Record As Variant
    
    With Purchase
        Record = Array( _
                    .WarehouseNumber, _
                    .Code, _
                    .PurchaseQuantity, _
                    .IsPickup, _
                    OPERATOR_CODE _
                    )
    End With
    
    Set TargetSheet = Worksheets("Magic�ꊇ�o�^")
    WriteRow = TargetSheet.UsedRange.Rows.Count + 1
    
    With TargetSheet
        .Cells(WriteRow, 2).NumberFormatLocal = String(9, "0")
        .Cells(WriteRow, 3).NumberFormatLocal = String(8, "0")
    
        .Cells(WriteRow, 1).Resize(1, 5).Value = Record
    End With
    
End Sub

Private Sub WriteMagicManualInput(ByRef Purchase As Purchase)

    Dim WriteRow As Long, TargetSheet As Worksheet, Record As Variant
    
    With Purchase
        Record = Array( _
                    .WarehouseNumber, _
                    .VendorCode, _
                    .VendorName, _
                    .Code, _
                    .ProductName, _
                    .PurchaseQuantity, _
                    .UnitCost, _
                    .IsPickup, _
                    OPERATOR_CODE _
                    )
    End With
    
    Set TargetSheet = Worksheets("Magic����͗p")
    WriteRow = TargetSheet.UsedRange.Rows.Count + 1
    
    TargetSheet.Cells(WriteRow, 4).NumberFormatLocal = "@"
    TargetSheet.Cells(WriteRow, 1).Resize(1, 9).Value = Record
    
End Sub
