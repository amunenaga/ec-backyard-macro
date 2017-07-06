Attribute VB_Name = "AppendHoldBook"
Option Explicit

Sub AppendHoldPurWokbook(ByVal HoldBook As Workbook)
'�����ۗ����X�g�ɁA�{���̎�z�ۗ����i��ǋL

With Worksheets(1)
    .Activate

    '�񐔂����킹��
    Columns("B").Insert
    Columns("B").ColumnWidth = 10
    Range("B1").Value = "���lA"
    
    Columns("D").Insert
    Columns("D").ColumnWidth = 10
    Range("D1").Value = "�A��"
    
    Columns("H").Insert
    Columns("H").ColumnWidth = 10
    Range("H1").Value = "���l"
    
    Columns("J").Insert
    Range("I:I").Copy Destination:=Range("J:J")
    

    Dim EndRow As Integer
    EndRow = .Range("A1").SpecialCells(xlCellTypeLastCell).Row
    
    '�����ۗ��V�[�g�̃R�s�[�͈͑I���AA�`M���2�s�ڂ���ŏI�s�͈͎̔擾
    Dim HoldProductRange As Range
    Set HoldProductRange = .Range("A2:M" & EndRow)
    
End With

'�����ۗ����J��
Dim HoldXlsxPath As String
HoldXlsxPath = "\\Server02\���i��\�l�b�g�̔��֘A\�����֘A\c�����ۗ���.xlsx"

Dim HoldLogWorkbook As Workbook
Set HoldLogWorkbook = FetchWorkBook(HoldXlsxPath)

'�ۗ��ꗗ�փR�s�[
With HoldLogWorkbook.Worksheets("�ۗ��ꗗ")

    '�t�B���^�[������
    If Not .AutoFilter Is Nothing Then
        Range("A1").AutoFilter
    End If
    
    '��s�폜�A�ۗ���Ɏ�z�����ۂɕۗ����X�g����f�[�^�ړ������邽�߁A��s�����邩��
    Call DeleteEmptyRow(HoldLogWorkbook.Worksheets("�ۗ��ꗗ"))
    
    '�ŏI�s�̓��t�`�F�b�N�A�ۗ��ꗗ�V�[�g�ł͕�����ŕێ����Ă��邽�߁A�����񓯎m�Ŕ�r����
    If CStr(.Range("G1").End(xlDown).Value) = Format(Date, "Mdd") Then
        HoldLogWorkbook.Close
        Exit Sub
    End If
    
    '�R�s�[���s
    Dim DestinationRange As Range
    Set DestinationRange = .Range("A1").End(xlDown).Offset(1, 0)
    
    HoldProductRange.Copy
    DestinationRange.PasteSpecial (xlPasteValues)
    
    Range("A1").Select
    
    HoldLogWorkbook.Save

End With

HoldLogWorkbook.Close

End Sub

Private Sub DeleteEmptyRow(HoldWorkSheet As Worksheet)
'�󔒍s���폜�A�s�𑖍��A�󔒍s�̃����W���擾����Range�I�u�W�F�N�g�̃��\�b�h�ł܂Ƃ߂č폜
'�Q�lURL  https://www.moug.net/tech/exvba/0050065.html

With HoldWorkSheet

    'UsedRange�v���p�e�B�Ȃ�A��s���܂߂čŏI�Z�����擾�ł���
    Dim UsedRowsCount As Long
    UsedRowsCount = .UsedRange.Rows.Count
    
    Dim i As Long, Target As Range

    '1��ڂ̃Z������Ȃ�ATarget�����W�Ƀ����W��ǉ����Ă���
    For i = 2 To UsedRowsCount
        
        Dim c As Range
        Set c = .Cells(i, 3)
        
        If c.Value = "" Then
            
            If Target Is Nothing Then
                Set Target = c.EntireRow
            Else
                Set Target = Union(Target, c.EntireRow)
            End If
            
        End If
    Next

    'Target�����W����s���ꊇ�폜
    If Not Target Is Nothing Then
        Target.Delete
    End If

End With

End Sub
