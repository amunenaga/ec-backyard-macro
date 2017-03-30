Attribute VB_Name = "DataValidate"
Option Explicit

Sub FilterLocation(Optional ByVal arg As Boolean)
'�A�h�C���Ŏ擾�����f�[�^���C������
'�G�N�Z���̃}�N���ꗗ�ɏo���Ȃ��悤�ɂ��邽�߈����t���Ƃ��Ă���B

OrderSheet.Activate

Dim LastRow As Long, i As Long
LastRow = Range("A1").SpecialCells(xlCellTypeLastCell).Row

For i = 2 To LastRow
    
    '���P�[�V�����C���A���i���o���f�[�V����
    Cells(i, 11).Value = CutOffUnlocation(Cells(i, 15).Value)
    
Next

End Sub

Function CutOffUnlocation(Location As String) As String
' ���K�\���Ń��P�[�V����[0-0-0-0][0- -0- - ][1-0-0-0-0]�Ȃǂ��폜���ĕԂ��܂��B

Dim Reg As New RegExp

Reg.Global = True
Reg.Pattern = "\[[0-9|\s]\-[0,1,2|\s]\-[0|\s]\-[0|\s]\-[0|\s]\]"

CutOffUnlocation = Reg.Replace(Location, "")

End Function

Function ValidateName(Name As String) As String

Dim Reg As New RegExp

Reg.Global = True
Reg.Pattern = ",|\!|\.|&"

Name = Reg.Replace(Name, "")


Reg.Pattern = "^((��|�y).*?(�z|��))*"
Name = Reg.Replace(Name, "")

ValidateName = Name

End Function
