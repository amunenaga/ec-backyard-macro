Attribute VB_Name = "SetParser"
Option Explicit

Const TIED_ITEM_LIST_BOOK As String = "��ď��iؽ�.xls"
Const LIST_BOOK_FOLDER As String = "\\server02\���i��\�l�b�g�̔��֘A\"
Sub ParseItems(r As Range)

'�Z�b�g���i���X�g�̃u�b�N���J��
Call OpenListBook

ThisWorkbook.Activate

Dim ComponentItems As Collection
Set ComponentItems = GetComponentItems(r.Value)

If ComponentItems Is Nothing Then Exit Sub

'�Z�b�g���e�����o������
Call InsertComponetRow(r, ComponentItems)

End Sub

Private Sub InsertComponetRow(c As Range, d As Collection)

Dim i As Long, v As Variant

For i = 1 To d.Count
    
    Rows(c.Offset(1, 0).Row).Insert (xlShiftDown)
    
    Set v = d.Item(i)
    
    '6�P�^�����6�P�^�A�Ȃ����JAN
    c.Offset(1, 0).NumberFormatLocal = "@"
    c.Offset(1, 0).Value = c.Value
    
    If v.Code <> "" Then

        c.Offset(1, 7).Value = v.Code
    
    Else
    
        c.Offset(1, 7).Value = v.Jan
    
    End If
    
    '���i����������
    c.Offset(1, 1).Value = v.Name
    
    '�P�̏��i�R�[�h�ƕK�v���ʏ�������
    c.Offset(1, 3).Value = v.Quantity * c.Offset(0, 3).Value
    c.Offset(1, 8).Value = v.Quantity * c.Offset(0, 3).Value
    
    '1�ڂ̃A�C�e���ɂ̂ݔ̔����i��t���ւ���
    '�����]�L�σt���O
    Dim Flg As Boolean
    
    If v.Quantity = 1 And Flg = False Then
    
        c.Offset(1, 2).Value = c.Offset(0, 2).Value
        c.Offset(0, 2).Value = 0
                
        Flg = True
        
    Else
        c.Offset(1, 2).Value = 0
        
    End If
    
    '�}����̍s�ɁA�󒍎��R�[�h�̓Z�b�g��7777�R�[�h������
    c.Offset(1, 0).Value = c.Value
    
    '�������A�}����̍s�ɒ����ԍ�������
    c.Offset(1, -1).Value = c.Offset(0, -1).Value
    
    c.Offset(1, 4).Value = c.Offset(0, 4).Value
    c.Offset(1, 5).Value = c.Offset(0, 5).Value
    c.Offset(1, 6).Value = c.Offset(0, 6).Value
    
Next

End Sub

Private Function GetComponentItems(TiedCode As String) As Collection
'�n���ꂽ�R�[�h����A�Z�b�g���eCollection��Ԃ��܂��B
'�Z�b�g���i���X�g�͌Ăяo�����̃v���V�[�W���ŊJ���Ă�����̂Ƃ��܂��B

'�Z�b�g���i���X�g����Y���R�[�h�̂���V�[�g�ƍs��T��

Dim i As Long
For i = 1 To Workbooks(TIED_ITEM_LIST_BOOK).Worksheets.Count
        
    Dim TiedCodeList As Worksheet
    Set TiedCodeList = Workbooks(TIED_ITEM_LIST_BOOK).Worksheets(i)

    Dim CodeRange As Range
    Set CodeRange = TiedCodeList.Range("A1:A" & TiedCodeList.Cells(2, 1).SpecialCells(xlCellTypeLastCell).Row)
        
    On Error Resume Next
        
        Dim HitRow As Double
        HitRow = WorksheetFunction.Match(TiedCode, CodeRange, 0)
        
        If HitRow > 0 Then Exit For
        
    On Error GoTo 0

Next

If HitRow = 0 Then
    Exit Function
End If

Dim ComponetItems As Collection
Set ComponetItems = New Collection

'E��=5����A�Z�b�g���e�̓X�^�[�g
'�w�b�_�[  SKU(�A��77777�n�܂�)/����(�ō�)/JAN�P�ʂ̑�����(���_���)     /JAN /����SKU /���� / ���i��

'��J�E���^
Dim k As Integer
k = TiedCodeList.Rows(1).Find("���i���1").Column

'IsEmpty���Ƌ󔒃Z���E���ꍇ������
Do Until TiedCodeList.Cells(HitRow, k) = ""

    Dim UnitCell As Range
    
    Dim Unit As ComponentItem
    Set Unit = New ComponentItem
    
    Set UnitCell = TiedCodeList.Cells(HitRow, k)
    
    With Unit
        
        .Jan = UnitCell.Value
        .Code = UnitCell.Offset(0, 1).Value
        .Name = UnitCell.Offset(0, 3).Value
        .Quantity = CLng(UnitCell.Offset(0, 2).Value)
    
    End With
        
    ComponetItems.Add Unit
    
    k = k + 4

Loop

Set GetComponentItems = ComponetItems

End Function

Private Sub OpenListBook()

'�Z�b�g���X�g�̃G�N�Z���t�@�C�����J�����A�J���Ă���΂��̂܂܏I�����܂��B
'1�̃s�b�L���O�V�[�g�̏����ŉ��񂩊J���ꍇ������̂ŁA����̂͌Ăяo�����ŃZ�b�g�����I���̃^�C�~���O�ōs���܂��B

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = TIED_ITEM_LIST_BOOK Then
        
        GoTo ret
    
    End If

Next

Set wb = Workbooks.Open(LIST_BOOK_FOLDER & TIED_ITEM_LIST_BOOK, ReadOnly:=True)

ret:

End Sub

Function CloseSetMasterBook(Optional ByVal arg As Variant) As Boolean

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = TIED_ITEM_LIST_BOOK Then
        
        wb.Close SaveChanges:=False
    
    End If

Next

End Function

Sub ParseScalingSet(r As Variant)

Dim Code As String, FixedCode As String
Code = r.Value

Dim SeparatedCode As Variant
SeparatedCode = Split(Code, "-", 2)

If SeparatedCode(0) Like String(5, "#") Then
    FixedCode = "0" & SeparatedCode(0)
Else
    FixedCode = SeparatedCode(0)
End If

'�P�̃R�[�h��I��ɓ����
Range("I" & r.Row).NumberFormatLocal = "@"
Range("I" & r.Row).Value = FixedCode

'IsNumeric���\�b�h�ŁA�n�C�t���̌�낪���l�ɕϊ��\���`�F�b�N
'�ϊ��\�Ȃ�A���Z�b�g�ƌ��Ȃ�

If Not IsNumeric(SeparatedCode(1)) Then
    Exit Sub
End If

'�Z�b�g�Ȃ�A�K�v���ʂ̓Z�b�g���ʁ~�󒍐��ʂɏ�������
Range("J" & r.Row).Value = Range("J" & r.Row).Value * CLng(Val(SeparatedCode(1)))


End Sub

