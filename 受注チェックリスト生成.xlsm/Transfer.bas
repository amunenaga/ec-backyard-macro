Attribute VB_Name = "Transfer"
Option Explicit

Sub ��ƃV�[�g�փf�[�^���o()

'�󒍃f�[�^�V�[�g���t�B���^�[���ĕK�v���ڂ̊y�V�����W�݂̂��R�s�[����

Sheet1.Range("A1").CurrentRegion.AutoFilter Field:=10, Criteria1:="�y�V�X"

Dim FilteredRange As Range
Set FilteredRange = Range("A1").CurrentRegion

'�K�v�ȗ�̃����W���w��
Dim RequireColumns As Range
Set RequireColumns = Columns("A:O")

Dim TargetRange As Range

'�t�B���^�[��̕K�v���ڗ�݂̂��R�s�[����
Intersect(FilteredRange, RequireColumns).Copy

'��ƃV�[�g�֓\��t���A�Z���̒���
'Paste���\�b�h�̎��s�_�C�A���O���o��ꍇ������̂ŁA���W���[���l�N�X�g�Ƃ���B
'�_�C�A���O�̕\���͍Č���������ł����A�N���b�v�{�[�h�̓��e�Ȃǂ��֌W���Ă���͗l�B
On Error Resume Next
With Worksheets.Add
    .Paste
    .Name = "��ƃV�[�g"
    
    '�񕝒��� ���i���A�͂���A�Z���͌Œ蕝 �P�ʁF�|�C���g
    .Columns("A:B").AutoFit
    .Columns("D:I").AutoFit
    .Columns("C").ColumnWidth = 40
    .Columns("K").ColumnWidth = 20
    .Columns("L").AutoFit
    .Columns("M:Q").ColumnWidth = 20
    
End With
On Error GoTo 0

'�󒍃f�[�^�V�[�g�̃I�[�g�t�B���^�[����
Sheet1.Range("A1").CurrentRegion.AutoFilter

'��̏����̂��߂ɁA��ɗ��}��
Columns("L").Insert
Range("L1").Value = "�͂���Z��"

Columns("C").Insert
Range("C1").Value = "JAN�R�[�h"

Range("A1").Select

End Sub
Sub �X�܎��ʔԍ��U��()

    
Worksheets("��ƃV�[�g").Activate

Dim i As Long
i = 2

Do
    '���[��������A�Г������p�̃��[���ԍ��֐U��ւ��āA�X�܃R�[�h����㏑��
    Dim Mall As String, MallId As Integer
    Mall = Cells(i, 11).Value
    
    Select Case Mall
        Case "Amazon�X"
            MallId = 1
        Case "�y�V�X"
            MallId = 2
        Case "Yahoo�X"
            MallId = 4
    End Select
    
    '�[�i���敪��DB�ł͐��l�^
    Cells(i, 10).NumberFormatLocal = "#"
    Cells(i, 10) = MallId
    
    i = i + 1

Loop Until IsEmpty(Cells(i, 1))

End Sub

Sub �Z������()
'�͂���s���{���A�͂���s�撬���A�͂���Z��1�A�͂���Z��2�A�͂���Z��3 �񂪕�����Ă���B
'�u�͂���Z���v��֌������Ċi�[�B

Worksheets("��ƃV�[�g").Activate

Dim i As Long
i = 2

Do
    'L��ɏZ��������
    Cells(i, 13).Value = Cells(i, 14).Value & Cells(i, 15).Value & Cells(i, 16).Value
    
    i = i + 1

Loop Until IsEmpty(Cells(i, 1))


End Sub

Sub JAN�]�L()
'���i�R�[�h��́A�����R�[�h �� �󔒂Ƃ��āA6�P�^�ȊO��JAN��֓]�L����

Worksheets("��ƃV�[�g").Activate

Dim i As Long
i = 2

Do
    Dim Code As String, Jan As String
    Code = Cells(i, 2).Value
    
    '����5�P�^��
    If Code Like String(6, "#") And InStr(1, Code, "0") = 1 Then
        
        Code = Right(Code, 5)
        Jan = ""
        
        Cells(i, 2).Resize(1, 2).Value = Array(Code, Jan)
    
    '5�P�^�ł�6�P�^�ł��Ȃ��ꍇ�AJAN��֓����
    ElseIf Not Code Like String(5, "#") And Not Code Like "5" & String(5, "#") Then
        
        Jan = Code
        Code = ""
    
        Cells(i, 2).Resize(1, 2).Value = Array(Code, Jan)
    
    End If

    i = i + 1

Loop Until IsEmpty(Cells(i, 1))

End Sub

Sub ���i���C��()

'���i������A�y�V�̃L�����y�[�������폜����
'��₩�y�z�Ő擪�ɋL�ڂ���Ă���̂ŁA���K�\���Ō��o���Ċ��ʂ��ƍ폜�A�������ʑΉ�
'�܂��ADB�̃t�B�[���h�T�C�Y��50�����Ȃ̂ŁA45�����ŃJ�b�g����B

Worksheets("��ƃV�[�g").Activate

'���[�v���Ŏg���s�J�E���^
Dim i As Long
i = 2

'���K�\���I�u�W�F�N�g�ƁA�p�^�[�����Z�b�g
Dim Reg As New RegExp
Reg.Global = True
Reg.Pattern = "^((��|�y).*?(�z|��))*"

Do
    Dim ProductName As String
    
    ProductName = Cells(i, 4).Value
    ProductName = Reg.Replace(ProductName, "")
            
    Cells(i, 4) = Left(ProductName, 45)
        
    i = i + 1

Loop Until IsEmpty(Cells(i, 1))

End Sub

Sub �����ƌ^�̕ύX()

Dim i As Long
i = 2

Do
    '�󒍔ԍ��̏C��
    Cells(i, 1).NumberFormatLocal = "#"
    Cells(i, 1).Value = CDbl(Cells(i, 1).Value)
    
    '���t�̕\�����C��
    Cells(i, 7).NumberFormatLocal = "yyyy/M/dd"
    Cells(i, 7).Value = Format(Cells(i, 7).Value, "yyyy/M/dd")
    
    i = i + 1

Loop Until IsEmpty(Cells(i, 1))

End Sub

Sub ��o�p�V�[�g�֓]�L()

Worksheets("��ƃV�[�g").Activate

'A2�`�ŏI�s�܂ŁA�Z�b�g���i�ȊO��]�L
Dim i As Long, k As Long
i = 2
k = 2

Do

    '7777�n�܂�͓]�L���Ȃ�
    If Cells(i, 3).Value Like "77777*" Then GoTo Continue
    
    '1�s���A���i�R�[�h�ƏZ�����R�s�[
    Dim Record As Range
    Set Record = Range(Cells(i, 1), Cells(i, 5))
    Set Record = Union(Record, Range(Cells(i, 7), Cells(i, 13)), Cells(i, 17))
    
    Record.Copy Worksheets("�A�b�v���[�h�V�[�g").Cells(k, 1)
    
    '�󒍖��׎}�Ԃ͑S��1�ł悢
    Worksheets("�A�b�v���[�h�V�[�g").Cells(k, 14).Value = "1"

    '�R�s�[��s�J�E���^���C���N�������g
    k = k + 1

Continue:
    i = i + 1
    
Loop Until IsEmpty(Cells(i, 1))

Worksheets("�A�b�v���[�h�V�[�g").Activate

End Sub

Function ValidateName(Name As String) As String

Dim Reg As New RegExp

Reg.Global = True
Reg.Pattern = "^((��|�y).*?(�z|��))*"
Name = Reg.Replace(Name, "")

ValidateName = Name

End Function

