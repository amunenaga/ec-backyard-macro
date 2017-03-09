Attribute VB_Name = "SetParser"
Option Explicit

'���[���o�^�R�[�h�̂����Z�b�g���i�ɂ��āA���e���i��6�P�^�EJAN�E���ʂ֕�������B

'Public Sub ParseTiedItems(CodeCell As Range)
'�s�̑}�����s�����߁A������Range�Ƃ���
'���ʃZ���̓R�[�h�̃Z������̃I�t�Z�b�g�Ŏ擾����


'Public Sub ParseMultipleSet(CodeCell As Range)
'�Z���ɓ����Ă��鐔�ʂ̏��������𔺂����߁A������Range
'���ʃZ���̓R�[�h�̃Z������̃I�t�Z�b�g�Ŏ擾����

'�萔 �Z�b�g���i���X�g�̃p�X
Const TIED_ITEM_LIST_BOOK As String = "��ď��iؽ�.xls"
Const LIST_BOOK_FOLDER As String = "\\server02\���i��\�l�b�g�̔��֘A\"

Sub �Z�b�g����()

Worksheets("��ƃV�[�g").Activate

Dim CodeRange As Range
Set CodeRange = Range(Cells(2, 2), Cells(Cells.SpecialCells(xlCellTypeLastCell).Row, 2))

Dim c As Range
For Each c In CodeRange

    Dim Code As String
    Code = c.Value
    
    '77777�n�܂�Z�b�g�R�[�h�Ȃ�
    If Code Like "77777*" Then
        
        Call ParseTiedItem(c)
    
    '-02 -04 -120 �n�C�t��-���� �Z�b�g�Ȃ� �n�C�t�����܂݃A���t�@�x�b�g�n�܂�łȂ�
    ElseIf InStr(Code, "-") > 1 And Not Code Like "[a-zA-Z]*" Then
        
        Call ParseMultipleSet(c)
    
    End If

Next

End Sub


Private Sub ParseTiedItem(CodeCell As Range)

Dim ComponentItems As Collection
Set ComponentItems = GetComponentItems(CodeCell.Value)

Dim i As Long, v As Variant

For i = 1 To ComponentItems.Count
    
    Set v = ComponentItems.Item(i)
    
    Rows(CodeCell.Offset(1, 0).Row).Insert (xlShiftDown)
    
    '6�P�^�����6�P�^�A�Ȃ����JAN
    CodeCell.Offset(1, 0).NumberFormatLocal = "@"
    
    If v.SyokonCode <> "" Then

        CodeCell.Offset(1, 0).Value = v.SyokonCode
    
    Else
    
        CodeCell.Offset(1, 0).Value = v.Jan
    
    End If
    
    '���i���o�͂ƕK�v���ʂ�������
    CodeCell.Offset(1, 2).Value = v.Name
    CodeCell.Offset(1, 3).Value = v.Quantity * CodeCell.Offset(0, 3).Value
    
    '�}����̍s�ɒ����ԍ�������
    CodeCell.Offset(1, -1).Value = CodeCell.Offset(0, -1).Value
    
    '�}����̍s��E��ȍ~�̒�����������
    CodeCell.Offset(1, 4).Resize(1, 12).Value = CodeCell.Offset(0, 4).Resize(1, 12).Value
    
Next

End Sub

Private Sub ParseMultipleSet(CodeCell As Range)
'012345-02�ȂǁA�n�C�t�� �����̃Z�b�g�𕪉����܂��B

'�R�[�h��������n�C�t���̈ʒu�ŕ���
Dim Code As String
Code = CodeCell.Value

Dim SeparatedCode As Variant
SeparatedCode = Split(Code, "-")

'�P�̃R�[�h����U�i�[
Dim ComponentCode As String
ComponentCode = SeparatedCode(0)

'IsNumeric���\�b�h�ŁA�n�C�t���̌��Ő��l�ɕϊ��\�Ȓl�����邩�`�F�b�N
'�ϊ��\�Ȃ��ڂ̐����Ɋ�Â��ā��Z�b�g�ƌ��Ȃ�
Dim i As Long

For i = 1 To UBound(SeparatedCode)

    If IsNumeric(SeparatedCode(i)) Then
        
        '�Z�b�g���ʂ��i�[
        Dim MultipleRatio As Long
        MultipleRatio = SeparatedCode(i)
        
        Exit For
    
    End If

Next

'������R�[�h�E���ʂ��o��

With CodeCell
    .NumberFormatLocal = "@"
    .Value = ComponentCode
End With

'���g�̐�������Z�ł��Ȃ��l����0�Ȃ̂ŁA�󒍐��ʂ����̂܂ܓ����B
'�o�א��ʂ̓Z�b�g���ʁ~�󒍐���

If MultipleRatio > 0 Then

    CodeCell.Offset(0, 3).Value = CodeCell.Offset(0, 3) * MultipleRatio

Else
    
    CodeCell.Offset(0, 3).Value = CodeCell.Offset(0, 3).Value
    
End If

End Sub

Private Function SearchTiedItemSheet(Code As String) As Worksheet
    '�u�Z�b�g�i���X�g�v�G�N�Z���t�@�C������A�Y���R�[�h�̂��郏�[�N�V�[�g��T���܂��B
    '���[�N�V�[�g����CountIf�ŊY���R�[�h�����邩�`�F�b�N�A����΂��̃��[�N�V�[�g��Ԃ��B
    
    Call OpenListBook
    
    Dim Hits As Long
    Dim i As Long
    
    For i = 1 To Workbooks(TIED_ITEM_LIST_BOOK).Worksheets.Count
        
        Dim LastRow As Long
        LastRow = Workbooks(TIED_ITEM_LIST_BOOK).Worksheets(i).Cells(2, 1).SpecialCells(xlCellTypeLastCell).Row
        
        Dim TiedItemCodeList As Range
        Set TiedItemCodeList = Workbooks(TIED_ITEM_LIST_BOOK).Worksheets(i).Range("A1:A" & LastRow)
        
        Hits = WorksheetFunction.CountIf(TiedItemCodeList, Code)

        If Hits > 0 Then
            
            Set SearchTiedItemSheet = Workbooks(TIED_ITEM_LIST_BOOK).Worksheets(i)
            Exit Function
        
        End If
    
    Next

End Function

Private Function GetComponentItems(TiedCode As String) As Collection

'�R�[�h����A�Z�b�g���e��

'�Z�b�g���i�R�[�h�̂���V�[�g��T��
Dim HitSheet As Worksheet
Set HitSheet = SearchTiedItemSheet(TiedCode)

'�o�^�R�[�h�̃����W�A������Match�֐��Œ��ׂāACode�̍s�ԍ����o��
Dim CodeRange As Range
Set CodeRange = HitSheet.Range("A1:A" & HitSheet.Cells(2, 1).SpecialCells(xlCellTypeLastCell).Row)

Dim HitRow As Double
HitRow = WorksheetFunction.Match(TiedCode, CodeRange, 0)

'���[�v���Ŏg���ϐ��ȂǏ�����
'�Z�b�g���e���i���i�[����R���N�V�������������B
'��J�E���^�A�z�񐔃J�E���^�A�Z�b�g���i�̏��i�����i�[����z��̏�����

Dim ComponetItems As Collection
Set ComponetItems = New Collection

Dim i As Integer
i = HitSheet.Rows(1).Find("���i���1").Column

'E��=5����A�Z�b�g���e�̓X�^�[�g
'�w�b�_�[  SKU(�A��77777�n�܂�)/����(�ō�)/JAN�P�ʂ̑�����(���_���)     /JAN /����SKU /���� / ���i��

'IsEmpty���Ƌ󔒃Z���E���ꍇ������
Do Until HitSheet.Cells(HitRow, i) = ""

    Dim UnitCell As Range, Unit As ComponentItem
    Set Unit = New ComponentItem
    
    Set UnitCell = HitSheet.Cells(HitRow, i)
            
    With Unit
        
        .Jan = UnitCell.Value
        .SyokonCode = UnitCell.Offset(0, 1).Value
        .Name = UnitCell.Offset(0, 3).Value
        .Quantity = CLng(UnitCell.Offset(0, 2).Value)
    
    End With
    
    ComponetItems.Add Unit
    
    i = i + 4

Loop

Set GetComponentItems = ComponetItems

End Function

Private Sub OpenListBook()

'�Z�b�g���X�g�̃G�N�Z���t�@�C�����J���܂��B1����Ɏ����ŕ��܂��B

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = TIED_ITEM_LIST_BOOK Then
        
        GoTo ret
    
    End If

Next

Set wb = Workbooks.Open(LIST_BOOK_FOLDER & TIED_ITEM_LIST_BOOK, ReadOnly:=True)

'�J�������[�N�u�b�N���A�N�e�B�u�ɂȂ�̂ŁA���̃u�b�N���A�N�e�B�u���������B
ThisWorkbook.Activate

ret:
'ret���x���ȉ��͖�����s����܂��B

Application.OnTime Now + TimeSerial(0, 1, 0), "CloseDataBook"

End Sub

Private Sub CloseDataBook()

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = TIED_ITEM_LIST_BOOK Then
        
        wb.Close SaveChanges:=False
    
    End If

Next

End Sub
