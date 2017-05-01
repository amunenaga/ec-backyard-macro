Attribute VB_Name = "DataValidate"
Option Explicit
Sub FixForAddin(Optional ByVal arg As Boolean)
'�Г�DB�Əƍ��ł���悤�Ɏ󒍃f�[�^�V�[�g�ɑ΂��āA�󒍏��i�R�[�h��̃R�[�h���A�h�C���p���i�R�[�h�֓]�L����B
'���g�Z�b�g�����������ōs���B

Worksheets("�󒍃f�[�^�V�[�g").Activate

Dim CodeRange As Range, c As Range
Set CodeRange = Range(Cells(2, 2), Cells(Range("B1").SpecialCells(xlCellTypeLastCell).Row, 2))

'�A�h�C���p�̃R�[�h���L������
For Each c In CodeRange
    
    Dim CurrentCodeCell As Range
    Set CurrentCodeCell = c
    
    'I��A�A�h�C�����s�p��6�P�^�������R�[�h�A��������JAN������
    Cells(c.Row, 9).NumberFormatLocal = "@"
    Cells(c.Row, 9).Value = DataValidate.ValidateCode(c.Value)
    
    '�K�v���ʁA��U�󒍂̐��ʂŖ��߂�B�Z�b�g������ɏ�����������B
    Cells(c.Row, 10).Value = Cells(c.Row, 4).Value

    '���g����
    If c.Value Like "*-*" Then
    
        Call SetParser.ParseScalingSet(c)
    
    End If

Next

End Sub
Sub FilterLocation(Optional ByVal arg As Boolean)
'�󒍃f�[�^�V�[�g�̑S�Ă̍s�ɑ΂��āA���P�[�V�����񂩂疳���ȃ��P�[�V�����������폜���ėL�����P�[�V������֓]�L�B

OrderSheet.Activate

Dim LastRow As Long, i As Long
LastRow = Range("A1").SpecialCells(xlCellTypeLastCell).Row

For i = 2 To LastRow
    
    '���P�[�V�����C���A���i���o���f�[�V����
    Cells(i, 11).Value = CutOffUnlocation(Cells(i, 15).Value)
    
Next

End Sub

Function CutOffUnlocation(Location As String) As String
'���K�\���Ń��P�[�V����[0-0-0-0][0- -0- - ][1-0-0-0-0]�Ȃǂ��폜���ĕԂ��܂��B

Dim Reg As New RegExp

Reg.Global = True

'���P�[�V�����̕��� �K-�ʘH-�I��-�i-��  �I�Ԃ�A�`Q�A���t�@�x�b�g
Reg.Pattern = "\[[0-9|\s]\-[0-2|\s]\-[0-9|\s]\-[0|\s]\-(([0-9]{1,})|\s)\]"

CutOffUnlocation = Reg.Replace(Location, "")

End Function

Function ValidateName(Name As String) As String
'���K�\���ŏ��i���̏C���B
'�J���}�E�s���I�h�Ȃǂ��폜�A�`���́y�z���Ŋ���ꂽ�y�V�̃Z�[�������폜


Dim Reg As New RegExp

Reg.Global = True
Reg.Pattern = ",|\!|\.|&"

Name = Reg.Replace(Name, "")


Reg.Pattern = "^((��|�y).*?(�z|��))*"
Name = Reg.Replace(Name, "")

ValidateName = Name

End Function

Function ValidateCode(Code As String) As String
'�R�[�h���󂯎���āA�����ȊO���폜�E13�P�^/6�P�^�ɑ���Ȃ��ꍇ�͖`��0��⊮�����R�[�h��Ԃ�

Dim FixedCode As String

'�A���t�@�x�b�g���폜
Dim Reg As New RegExp
Reg.Global = True
Reg.Pattern = "[a-zA-Z]"
Code = Reg.Replace(Code, "")

'6�P�^�Ȃ炻�̂܂ܓ����
If Code Like String(6, "#") Then
    FixedCode = Code

'����5�P�^�͓��Ƀ[����ǋL
ElseIf Code Like String(5, "#") Then
    
    FixedCode = "0" & Code

'JAN�����̂܂ܓ����
ElseIf Code Like String(13, "#") Then
    
    FixedCode = Code
    
'����7�P�^�ȏ�12�P�^�Ȃ�A13�P�^�ɂȂ�悤�擪��0��ǋL
ElseIf Code Like (String(7, "#") & "*") And Len(Code) <= 12 Then

    FixedCode = String(13 - Len(Code), "0") & Code
    
Else
'�ǂ̏����ɂ���v���Ȃ��ꍇ�ł��A�l�͕Ԃ�
    
    FixedCode = Code
    
End If

ValidateCode = FixedCode

End Function
