Attribute VB_Name = "Sort"
Option Explicit

Sub �U���p�V�[�g_�\�[�g(Sheet As Worksheet)
Attribute �U���p�V�[�g_�\�[�g.VB_ProcData.VB_Invoke_Func = " \n14"

Dim SortRange As Range
Set SortRange = Sheet.Range("A1").CurrentRegion

Dim CodeRange As Range
Set CodeRange = Sheet.Range("A2:A" & SortRange.Rows.Count)

'�\�[�g�������Z�b�g
With Sheet.Sort
    
    '��U�\�[�g���N���A
    .SortFields.Clear
    
    '�\�[�g�L�[���Z�b�g ���L�[ ���i�R�[�h�F�F�A���L�[ ���i�R�[�h�F����
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

'�J�����g���[�W�������Z���N�g����Ă���̂ŁA�I���Z�����Z��A1�ɃZ�b�g������
Sheet.Activate
Range("A1").Select

End Sub
