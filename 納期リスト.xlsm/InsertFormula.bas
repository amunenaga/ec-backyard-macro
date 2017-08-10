Attribute VB_Name = "InsertFormula"
Option Explicit

Sub ���ד��̎Z�o�������()
'�Z�o���������͂̍s�𒲂ׂ�
'A�񂩂�K��̏��i����܂ł́A��z�f�[�^�쐬������͂����

ThisWorkbook.Worksheets("�[�����X�g").Activate

Dim InsertEndCell As Range, InsertStartCell As Range, TargetRange As Range

'�ŏI�s��L��ɐ����������Ă���΁A���t�Z�o���͓��͍ςȂ̂łƂ��āA�v���V�[�W���I��
If Cells(Range("A1").End(xlDown).Row, 12).Formula <> "" Then Exit Sub

'�ŏI�s��H��̃Z��
Set InsertEndCell = Cells(Range("A1").End(xlDown).Row, 8)

'L��Ŏ��������Ă���ŏI�s�����s����H��̃Z��
Set InsertStartCell = Cells(InsertEndCell.Row, 12).End(xlUp).Offset(1, -4)

Set TargetRange = Range(InsertStartCell, InsertEndCell)

'�Z�b�g���������W�ɑ΂��Ď�������
Dim r As Variant

For Each r In TargetRange

    'H��AI��̓��[�J�[�V�[�g�̓��ׂɊւ��镶��
    On Error Resume Next
        r.Offset(0, 0).Value = WorksheetFunction.VLookup(Cells(r.Row, 3).Value, Worksheets("���[�J�[").Range("B3:D1000"), 2, False)
        r.Offset(0, 1).Value = WorksheetFunction.VLookup(Cells(r.Row, 3).Value, Worksheets("���[�J�[").Range("B3:D1000"), 3, False)
    On Error GoTo 0

    'K , L��ͤW��̓��t������ד����Z�o���鎮������A�s�ԍ������ƒu���Ēu�����Ď��̕���������
    r.Offset(0, 3).Formula = Replace("=IFERROR(VALUE(IF($J@="""",IF($I@=""����"",$E@,IF($I@=""����"",$E@+1,IF($I@=""���X��"",$E@+2,""""))),$J@+1)),"""")", "@", r.Row)
    r.Offset(0, 4).Formula = Replace("=IF($K@="""","""",IF(MOD($K@,7)=0,$K@+2,IF(MOD($K@,7)=1,$K@+1,$K@)))", "@", r.Row)

Next

End Sub
