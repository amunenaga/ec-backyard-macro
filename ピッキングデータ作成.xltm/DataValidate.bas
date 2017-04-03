Attribute VB_Name = "DataValidate"
Option Explicit
Sub FixForAddin()

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

Function ValidateCode(Code As String) As String

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
