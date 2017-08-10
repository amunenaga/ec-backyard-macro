Attribute VB_Name = "SheetMaintenance"
Option Explicit

Sub �����t�������͈͏C��()
'�\���̍X�V�������ݒ肳�ꂽ�͈͑S�Ăɑ���̂ŁA�ŏ����̍ŏI2�T�ԕ��ɂ��Ă����B

'���X�̏����ł͎�z�f�[�^�쐬����ǋL����Ƃ��Ɏ��s

'@URL:https://msdn.microsoft.com/ja-jp/library/office/ff837422.aspx (Excell2013�ȍ~�Ƃ��邪�A2010�ł������j
'@URL:http://www.relief.jp/docs/excel-vba-check-if-has-format-conditions.html

With Worksheets("�[�����X�g")

    Dim WeekBeforeLastRow As Long, Endrow As Long, LatestArrivalRange
    Endrow = .Range("E2").End(xlDown).Row
    
    'T���14���O�̓��t�̍s���擾�A�y���͎�z���͂��Ȃ����ߌ��������́u�ȉ��v�ŃZ�b�g�A���t�̌������̓_�u���^�ŃV���A��������
    WeekBeforeLastRow = WorksheetFunction.Match(CDbl(DateAdd("d", -14, Date)), .Range(Cells(1, 5), Cells(Endrow, 5)), 1) + 1

    Set LatestArrivalRange = Range(Cells(WeekBeforeLastRow, 10), Cells(Endrow, 10))

    .Cells.FormatConditions(1).ModifyAppliesToRange LatestArrivalRange

End With

End Sub

Sub �ꃖ���ȑO�]�L()
'��z�f�[�^�쐬����ǋL����Ƃ��Ɏ��s

'���X�̏����ł͎�z�f�[�^�쐬����ǋL����Ƃ��Ɏ��s

'ETA; Estimated time of arrival ��s�@�̓����\�莞���̂���
Dim Endrow As Long, EtaSheet As Worksheet, LogSheet As Worksheet, i As Long

Set EtaSheet = ThisWorkbook.Worksheets("�[�����X�g")
Set LogSheet = ThisWorkbook.Worksheets("��")

EtaSheet.Activate

Endrow = EtaSheet.Range("F2").End(xlDown).Row

'�ŏ��̃f�[�^��1�����ȏ�O�Ȃ�A���V�[�g�փR�s�[���邽�߂�2�s�ڂ��R�s�[�Ώ۔͈͂Ƃ��Ă��Z�b�g
If DateDiff("d", Date, Cells(2, 6).Value) < -30 Then
    Dim TargetRange As Range
    Set TargetRange = Range("A2:AA2")
Else
    Exit Sub
End If

'1�����ȏ�O�́A�[�����X�g��3�s�ڂ���Range��A�����Ă���
i = 2

Do

    Set TargetRange = Union(TargetRange, EtaSheet.Range(Cells(i, 1), Cells(i, 27)))
    i = i + 1

Loop While DateDiff("d", Date, Cells(i, 6).Value) < -30

'1�����ȏ�O�͈̔͂��R�s�[���č폜
TargetRange.Copy
LogSheet.Cells(LogSheet.UsedRange.Rows.Count + 1, 1).PasteSpecial Paste:=xlPasteValues

TargetRange.EntireRow.Delete shift:=xlShiftUp

End Sub
