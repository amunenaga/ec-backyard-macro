Attribute VB_Name = "Output"
Option Explicit

Sub QtyCsv()
'FileSystemObject�̃e�L�X�g�X�g���[����CSV�t�@�C���𐶐����āATextStream�œ��e�𗬂����݂܂��B
'���b�ŏI���܂��B

With yahoo6digit '���t�[�f�[�^�̉�����

    .Activate
    
    '�u"�o�^�Ȃ�"�v�Ɓu"��"�v ����2�ȊO���t�B���^�[�ŕ\���cTODO�F1��ڂ���t�B���^�[�̏󋵂��`�F�b�N��������������
    '16-2-29 �p�Ԃ̋敪���u���p�ԁv�ɂȂ�܂����B
    
    Dim StatusCol As Integer
    StatusCol = .Rows(1).Find("status").Column
    
    .Range("A1").CurrentRegion.AutoFilter Field:=StatusCol, Criteria1:=Array( _
            "�r�o����", "����i", "�݌ɔp��", "�݌ɏ���", "�I�Ȃ��ɗL", "�I�Ȃ�����", "��������", "�o�^�̂�", "���p�ԕi", "�̘H����", "�̔����~", "�W��" _
            ), Operator:=xlFilterValues
    
    '�t�B���^�[���������W���Z�b�g�ACSV�̃w�b�_�[�͕ʓr��������ł����̂ŁA2�s�ڈȍ~�̃����W�B
    Dim A As Range
    Set A = .Range("C1").CurrentRegion.SpecialCells(xlCellTypeVisible)
    
    Dim B As Range
    Set B = .Range("C2").Resize(Range("C1").SpecialCells(xlCellTypeLastCell).Row - 1, 1)
    
    Dim CodeRange As Range
    Set CodeRange = Application.Intersect(A, B)

End With

'�����o���pCSV��p��
Dim day As String
day = Format(Date, "mm") & Format(Date, "dd")

Dim OutputCsvName As String
OutputCsvName = "�����݌ɃA�b�v�p" & day & ".csv"

Dim FSO As Object 'TODO:���O�o�C���f�B���O�ɕύX
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim TS As Object
    
Set TS = FSO.CreateTextFile(Filename:=ThisWorkbook.Path & "\" & OutputCsvName, _
                            OverWrite:=True)
                            
'�w�b�_�[����������
Dim header As Variant
header = "code,quantity,allow-overdraft"

TS.WriteLine header

Dim colQuantity As Long, colAllow As Long, colStatus As Long

colQuantity = yahoo6digit.Rows(1).Find("quantity").Column
colAllow = yahoo6digit.Rows(1).Find("allow-overdraft").Column
colStatus = yahoo6digit.Rows(1).Find("status").Column

'�R�[�h�����W�ɑ΂��āAr.row�ōs�ԍ������o���ē����s��Quantity/Allow�̒l���擾����
Dim r As Range, Code As String, Qty As String, Pur As String

For Each r In CodeRange
    
    Code = r.Value
    
    Qty = Cells(r.Row, colQuantity).Value
    Pur = Cells(r.Row, colAllow).Value
    
    TS.WriteLine Code & "," & Qty & "," & Pur

Next

TS.Close

End Sub

