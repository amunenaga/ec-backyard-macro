Attribute VB_Name = "Transfer"
Option Explicit

Sub ��ƃV�[�g�փf�[�^���o()

'�Y���V�[�g���t�B���^�[���ĕK�v��̃����W�݂̂��R�s�[����

Sheet1.Range("A1").CurrentRegion.AutoFilter Field:=2, Criteria1:="<>"

Dim FilteredRange As Range
Set FilteredRange = Range("A1").CurrentRegion

'�K�v�ȗ�̃����W���w��
Dim RequireColumns As Range
Set RequireColumns = Columns("A:P")

Dim TargetRange As Range

'�t�B���^�[��̕K�v���ڗ�݂̂��R�s�[����
Intersect(FilteredRange, RequireColumns).Copy

'��ƃV�[�g�֓\��t���A�Z���̒���
With Worksheets.Add
    .Paste
    .Name = "��ƃV�[�g"
    
    '�񕝒��� ���i���A�͂���A�Z���͌Œ蕝 �P�ʁF�|�C���g
    .Columns("A:B").AutoFit
    .Columns("D:I").AutoFit
    .Columns("C").ColumnWidth = 40
    .Columns("K").ColumnWidth = 20
    .Columns("L").AutoFit
    .Columns("M:P").ColumnWidth = 20
    
End With

'�I�[�g�t�B���^�[����
Sheet1.Range("A1").CurrentRegion.AutoFilter

'��̏����̂��߂ɁA��ɗ��}��
Columns("L").Insert
Range("L1").Value = "�͂���Z��"

Columns("C").Insert
Range("C1").Value = "JAN�R�[�h"

Range("A1").Select

End Sub

Sub �Z������()
'�͂���s���{���A�͂���s�撬���A�͂���Z��1�A�͂���Z��2�A�͂���Z��3 �񂪕�����Ă���B
'�u�͂���Z���v��֌������Ċi�[�B

Worksheets("��ƃV�[�g").Activate

Dim i As Long
i = 2

Do
    'L��ɏZ��������
    Cells(i, 13).Value = Cells(i, 14).Value & Cells(i, 15).Value & Cells(i, 16).Value & Cells(i, 17).Value & Cells(i, 18).Value
    
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

Sub �y�V���i���C��()

'���i������A�y�V�̃L�����y�[�������폜����
'��₩�y�z�Ő擪�ɋL�ڂ���Ă���̂ŁA�����Instr�� �� �z�̈ʒu�����o���đO���폜
'���K�\���ł̎����́Awhatnot�󒍎捞�}�N���ɂ���B

Worksheets("��ƃV�[�g").Activate

'�L�����y�[�������͂��Ă�����J�b�R�z����`
Dim CloseBrackets() As Variant
CloseBrackets = Array(Array("�y", "�z"), Array("��", "��"))

'�s�J�E���^
Dim i As Long
i = 2

Do
    Dim ProductName As String
    ProductName = Cells(i, 3).Value
    
    '�����ʂ��������ڂɏo�Ă��邩���ׂ�
    Dim k As Integer
    For k = 0 To UBound(CloseBrackets)
    
        '�L�����y�[�����̊��ʂ��`���ɂ��邩�`�F�b�N
        If InStr(1, ProductName, CloseBrackets(k)(0)) = 1 Then
            
            Dim CampaignInfoCharEnd As Integer
            CampaignInfoCharEnd = InStr(1, ProductName, CloseBrackets(k)(1)) + 1
            
            Debug.Assert CampaignInfoCharEnd = 0
            
            Cells(i, 3) = Mid(ProductName, CampaignInfoCharEnd)
            
        End If
        
    Next
        
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
    If Cells(i, 3).Value Like "77777" Then GoTo Continue
    
    '1�s���̃����W���`
    Dim Record As Range
    Set Record = Range(Cells(i, 1), Cells(i, 5))
    
    Set Record = Union(Record, Range(Cells(i, 7), Cells(i, 13)))
    
    '�s���R�s�[�A�R�s�[��s�J�E���^���C���N�������g
    Record.Copy Worksheets("��o�V�[�g").Cells(k, 1)
    k = k + 1

Continue:
    i = i + 1
    
Loop Until IsEmpty(Cells(i, 1))

Worksheets("��o�V�[�g").Activate

End Sub
