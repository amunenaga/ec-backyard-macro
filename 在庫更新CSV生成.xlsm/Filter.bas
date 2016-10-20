Attribute VB_Name = "Filter"
Option Explicit


Sub SetStatusFilter()

With yahoo6digit
    
    .Activate
       
    Dim StatusCol As Integer
    StatusCol = .Rows(1).Find("status").Column
    
    '"�r�o����" �ȉ��̕�����́A�����́u�敪�v�ɍ��킹�Ă�������
    '�󔒂Ɠo�^�Ȃ����t�B���^�[�Ŕ�\���ɂ��܂��A�}�N���Ńt�B���^�[���L�^���ď���������Ɗy�ł��B
    
    .Range("A1").CurrentRegion.AutoFilter Field:=StatusCol, Criteria1:=Array( _
    "�r�o����", "���p�ԕi", "����i", "�݌ɏ���", "�݌ɔp��", "��������", "�o�^�̂�", "�̔����~", "�̘H����", "�W��"), Operator _
    :=xlFilterValues

End With

End Sub
