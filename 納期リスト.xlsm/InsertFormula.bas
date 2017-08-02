Attribute VB_Name = "InsertFormula"
Option Explicit

Sub ���ד��̎Z�o�������()
'�Z�o���������͂̍s�𒲂ׂ�
'A�񂩂�K��̏��i����܂ł́A��z�f�[�^�쐬������͂����

ThisWorkbook.Worksheets("�[�����X�g").Activate

Dim InsertEndCell As Range, InsertStartCell As Range, TargetRange As Range
'�ŏI�s��U��̃Z��
Set InsertEndCell = Cells(Range("A1").End(xlDown).Row, 21)
'X��Ŏ��������Ă���ŏI�s�����s����U��̃Z��
Set InsertStartCell = Cells(InsertEndCell.Row, 24).End(xlUp).Offset(1, -3)

Set TargetRange = Range(InsertStartCell, InsertEndCell)

'�Z�b�g���������W�ɑ΂��Ď�������
Dim r As Variant

For Each r In TargetRange

    'U��AV��̓��[�J�[�V�[�g�̓��ׂɊւ��镶��
    On Error Resume Next
        r.Offset(0, 0).Value = WorksheetFunction.VLookup(Cells(r.Row, 4).Value, Worksheets("���[�J�[").Range("B3:D1000"), 2, False)
        r.Offset(0, 1).Value = WorksheetFunction.VLookup(Cells(r.Row, 4).Value, Worksheets("���[�J�[").Range("B3:D1000"), 3, False)
    On Error GoTo 0

    'X , Y��ͤW��̓��t������ד����Z�o���鎮������A�s�ԍ������ƒu���Ēu�����Ď��̕���������
    r.Offset(0, 3).Formula = Replace("=IFERROR(VALUE(IF($W@="""",IF($V@=""����"",$F@,IF($V@=""����"",$F@+1,IF($V@=""���X��"",$T@+2,""""))),$W@+1)),"""")", "@", r.Row)
    r.Offset(0, 4).Formula = Replace("=IF($X@="""","""",IF(MOD($X@,7)=0,$X@+2,IF(MOD($X@,7)=1,$X@+1,$X@)))", "@", r.Row)

Next

End Sub
