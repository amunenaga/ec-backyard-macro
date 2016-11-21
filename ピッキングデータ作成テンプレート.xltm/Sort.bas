Attribute VB_Name = "Sort"
Option Explicit

Sub �U���p�V�[�g_�\�[�g()
Attribute �U���p�V�[�g_�\�[�g.VB_ProcData.VB_Invoke_Func = " \n14"
    
Dim SortRange As Range
Set SortRange = Worksheets("�U�����p�ꗗ�V�[�g").Range("A1").CurrentRegion

Dim CodeRange As Range
Set CodeRange = Worksheets("�U�����p�ꗗ�V�[�g").Range("C2:C" & SortRange.Rows.Count)

'�\�[�g�������Z�b�g
With Worksheets("�U�����p�ꗗ�V�[�g").Sort
    
    '��U�\�[�g���N���A
    .SortFields.Clear
    
    '�\�[�g�L�[���Z�b�g ���L�[ ���i�R�[�h�F�F�A���L�[ ���i�R�[�h�F����
    '�Z���w�i�ɐF�����Ă���I���������Ɍł߂�B�w�i�Ȃ����I����A�w�i�F�����I���� ���ꂼ��̒���6�P�^����
    .SortFields.Add Key:=CodeRange, SortOn:=xlSortOnCellColor, Order:=xlAscending, DataOption:=xlSortNormal
    .SortFields.Add Key:=CodeRange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    '�\�[�g�Ώۂ̃f�[�^�������Ă�͈͂��w�肵��
    .SetRange SortRange
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    
    '�Z�b�g����������K�p
    .Apply

End With

End Sub
