Attribute VB_Name = "DataVaridate"
Option Explicit

Function ModifyOrderSheet(Optional var As Variant) As Boolean
'�A�h�C���Ŏ擾�����f�[�^��]�L

OrderSheet.Activate

Dim LastRow As Long, i As Long
LastRow = Range("A1").SpecialCells(xlCellTypeLastCell).Row

For i = 2 To LastRow
    
    '���P�[�V�����C��
    Cells(i, 11).Value = CutOffUnlocation(Cells(i, 15).Value)

Next

End Function

Private Function CutOffUnlocation(Location As String) As String
' ���K�\���Ń��P�[�V����[0-0-0-0][0- -0- - ][1-0-0-0-0]�Ȃǂ��폜���ĕԂ��܂��B

Dim Reg As New RegExp

Reg.Global = True
Reg.Pattern = "\[\d00:[0|\s]\-[0|\s]\-[0|\s]\-[0|\s]\-[0|\s]\]"

CutOffUnlocation = Reg.Replace(Location, "")

End Function

Private Function CutCampaignWord(Name As String) As String



End Function


