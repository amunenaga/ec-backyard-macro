Attribute VB_Name = "AppendRefaxLog"
Option Explicit

Sub AppendRefaxList()
'FAX�[���񓚃��X�g�ɖ{���̔������i�E�����ۗ���ǋL����

'FAX�[���񓚃��X�g���J��
Dim RefaxBook As Workbook, RefaxSheet As Worksheet, WriteCell As Range
Set RefaxBook = FetchWorkBook("\\Server02\���i��\�l�b�g�̔��֘A\�����֘A\�����������o�b�N�A�b�v\FAX�[���񓚃��X�g.xlsm")
Set RefaxSheet = RefaxBook.Worksheets("�[�����X�g")

RefaxSheet.Activate

If DateDiff("d", Date, Cells(Range("A1").CurrentRegion.Rows.Count, 5).Value) >= 0 Then
    RefaxBook.Close SaveChanges:=False
    Exit Sub
End If
'�ԐMFAX�̍ŏI�̋󔒍s�֏�������
Set WriteCell = RefaxSheet.Cells(Range("A1").CurrentRegion.Rows.Count, 1).Offset(1, 0)

'�������i���X�g����f�[�^���R�s�[
Dim DataCol As Range, PurchaseSheet As Worksheet
ThisWorkbook.Worksheets("�������i���X�g").Activate
ThisWorkbook.Worksheets("�������i���X�g").Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

If Range("A2").Value = "" Then
    RefaxBook.Close SaveChanges:=True
    Exit Sub
Else
    Set DataCol = ThisWorkbook.Worksheets("�������i���X�g").Range(Cells(2, 1), Cells(2, 1).End(xlDown))
End If

'DataCol�����W��WriteCell���E�փI�t�Z�b�g���Ȃ���f�[�^���R�s�[���Ă����B

DataCol.Offset(0, 6).Copy Destination:=WriteCell '��������
DataCol.Offset(0, 0).Copy Destination:=WriteCell.Offset(0, 1) '����
DataCol.Offset(0, 1).Copy Destination:=WriteCell.Offset(0, 2) '�d����
DataCol.Offset(0, 2).Copy Destination:=WriteCell.Offset(0, 3) '���[�����ʋL��
DataCol.Offset(0, 3).Copy Destination:=WriteCell.Offset(0, 4) '���t
DataCol.Offset(0, 4).Copy Destination:=WriteCell.Offset(0, 5) '���i�R�[�h
DataCol.Offset(0, 5).Copy Destination:=WriteCell.Offset(0, 6) '���i��


'���l�ɕۗ��V�[�g����f�[�^���R�s�[

Dim HoldSheet As Worksheet
ThisWorkbook.Worksheets("�ۗ�").Activate

Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

If Range("A2").Value <> "" Then
    Set DataCol = Worksheets("�ۗ�").Range(Cells(2, 1), Cells(1, 1).End(xlDown))
    
    RefaxSheet.Activate
    Set WriteCell = RefaxSheet.Cells(Range("A1").CurrentRegion.Rows.Count, 1).Offset(1, 0)
    
    '���ʂ̓��Ɂu�ۗ��v���������ē\��t��
    Dim HoldQty As Variant, i As Long
    'HoldQty�͓񎟌��z��Ŋi�[�����
    HoldQty = DataCol.Offset(0, 6).Value
    
    '�ۗ���1�s�̎��́A�z��ɂȂ�Ȃ�
    If IsArray(HoldQty) Then
        
        For i = 1 To UBound(HoldQty)
            HoldQty(i, 1) = "�ۗ��F" & HoldQty(i, 1)
        Next
        WriteCell.Resize(UBound(HoldQty), 1).Value = HoldQty
    
    Else
    
        WriteCell.Value = "�ۗ��F" & HoldQty
    
    End If
    
    DataCol.Offset(0, 1).Copy Destination:=WriteCell.Offset(0, 1) '����
    DataCol.Offset(0, 2).Copy Destination:=WriteCell.Offset(0, 2) '�d����
    DataCol.Offset(0, 3).Copy Destination:=WriteCell.Offset(0, 3) '���[�����ʋL��
    DataCol.Offset(0, 4).Copy Destination:=WriteCell.Offset(0, 4) '���t
    DataCol.Offset(0, 5).Copy Destination:=WriteCell.Offset(0, 5) '��z�����i�R�[�h
    DataCol.Offset(0, 7).Copy Destination:=WriteCell.Offset(0, 6) '���i��
    
    DataCol.Offset(0, 0).Copy
    WriteCell.Offset(0, 9).PasteSpecial Paste:=xlPasteValues   '�ۗ����R

End If

'FAX�[���񓚃��X�g.xlsm�̃}�N�������s����
On Error Resume Next
    Application.Run "FAX�[���񓚃��X�g.xlsm!�ꃖ���ȑO�]�L"
    Application.Run "FAX�[���񓚃��X�g.xlsm!�O�����ȑO�폜"
    Application.Run "FAX�[���񓚃��X�g.xlsm!���ד��̎Z�o�������"
    Application.Run "FAX�[���񓚃��X�g.xlsm!�����t�������͈͏C��"
On Error GoTo 0

RefaxBook.Save

End Sub


