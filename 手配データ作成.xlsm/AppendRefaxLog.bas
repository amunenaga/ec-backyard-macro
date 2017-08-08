Attribute VB_Name = "AppendRefaxLog"
Option Explicit

Sub AppendRefaxList()
'FAX�[���񓚃��X�g�ɖ{���̔������i�E�����ۗ���ǋL����

'FAX�[���񓚃��X�g���J��
Dim RefaxBook As Workbook, RefaxSheet As Worksheet, WriteCell As Range
Set RefaxBook = FetchWorkBook("\\Server02\���i��\�l�b�g�̔��֘A\�����֘A\�����������o�b�N�A�b�v\FAX�[���񓚃��X�g.xlsm")
Set RefaxSheet = RefaxBook.Worksheets("�[�����X�g")

RefaxSheet.Activate

If DateDiff("d", Date, Cells(Range("A1").CurrentRegion.Rows.Count, 6).Value) >= 0 Then
    RefaxBook.Close SaveChanges:=False
    Exit Sub
End If
'�ԐMFAX�̍ŏI�̋󔒍s�֏�������
Set WriteCell = Cells(Range("A1").CurrentRegion.Rows.Count, 1).Offset(1, 0)

'�������i���X�g����f�[�^���R�s�[
Dim DataCol As Range
ThisWorkbook.Worksheets("�������i���X�g").Activate
Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

If Range("A2").Value = "" Then
    RefaxBook.Close SaveChanges:=True
    Exit Sub
Else
    Set DataCol = Range(Cells(2, 1), Cells(2, 1).End(xlDown))
End If

'DataCol�����W��WriteCell���E�փI�t�Z�b�g���Ȃ���f�[�^���R�s�[���Ă����B

DataCol.Offset(0, 6).Copy Destination:=WriteCell '��������
DataCol.Offset(0, 0).Copy Destination:=WriteCell.Offset(0, 1) '����
DataCol.Offset(0, 1).Copy Destination:=WriteCell.Offset(0, 3) '�d����
DataCol.Offset(0, 2).Copy Destination:=WriteCell.Offset(0, 4) '���[�����ʋL��
DataCol.Offset(0, 3).Copy Destination:=WriteCell.Offset(0, 5) '���t
DataCol.Offset(0, 4).Copy Destination:=WriteCell.Offset(0, 8) '���i�R�[�h
DataCol.Offset(0, 5).Copy Destination:=WriteCell.Offset(0, 10) '���i��


'���l�ɕۗ��V�[�g����f�[�^���R�s�[

ThisWorkbook.Worksheets("�ۗ�").Activate
Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

If Range("A2").Value <> "" Then
    Set DataCol = Range(Cells(2, 1), Cells(2, 1).End(xlDown))
    
    RefaxSheet.Activate
    Set WriteCell = Cells(Range("A1").CurrentRegion.Rows.Count, 1).Offset(1, 0)
    
    '���ʂ̓��Ɂu�ۗ��v���������ē\��t��
    Dim HoldQty As Variant, i As Long
    HoldQty = DataCol.Offset(0, 6).Value
    
    For i = 1 To UBound(HoldQty)
        HoldQty(i, 1) = "�ۗ��F" & HoldQty(i, 1)
    Next
    
    WriteCell.Resize(UBound(HoldQty), 1).Value = HoldQty
    
    DataCol.Offset(0, 1).Copy Destination:=WriteCell.Offset(0, 1) '����
    DataCol.Offset(0, 2).Copy Destination:=WriteCell.Offset(0, 3) '�d����
    DataCol.Offset(0, 3).Copy Destination:=WriteCell.Offset(0, 4) '���[�����ʋL��
    DataCol.Offset(0, 4).Copy Destination:=WriteCell.Offset(0, 5) '���t
    DataCol.Offset(0, 5).Copy Destination:=WriteCell.Offset(0, 8) '��z�����i�R�[�h
    DataCol.Offset(0, 7).Copy Destination:=WriteCell.Offset(0, 10) '���i��
    
    DataCol.Offset(0, 0).Copy
    WriteCell.Offset(0, 22).PasteSpecial Paste:=xlPasteValues   '�ۗ����R

End If

On Error Resume Next
    Application.Run "FAX�[���񓚃��X�g.xlsm!�ꃖ���ȑO�]�L"
    Application.Run "FAX�[���񓚃��X�g.xlsm!���ד��̎Z�o�������"
    Application.Run "FAX�[���񓚃��X�g.xlsm!�����t�������͈͏C��"
On Error GoTo 0

RefaxBook.Save

End Sub


