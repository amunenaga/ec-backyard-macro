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

Private Sub InsertComponetRow(c As Range, ComponentItems As Collection)

Dim i As Long
For i = 1 To ComponentItems.Count
    
    Dim Record As Range
    Set Record = Range(Cells(c.Row, 1), Cells(c.Row, 15))
    
    '��U�Z�b�g�i�̍s���R�s�[
    Record.Copy
    Record.Offset(1, 0).Insert (xlShiftDown)
    
    '�}�������s�������s�ԍ�
    Dim wr As Long
    wr = c.Row + 1
    
    '�}����̍s���Z�b�g���e�̏��i���ŏ���������
    Dim Component As Variant
    Set Component = ComponentItems.Item(i)
    
    '�A�h�C���p�̃R�[�h��6�P�^�����6�P�^�A�Ȃ����JAN�ŏ㏑��
    If Component.Code <> "" Then

        Cells(wr, 9).Value = Component.Code
    
    Else
    
        Cells(wr, 9).Value = Component.Jan
    
    End If
    
    '���i���㏑��
    Cells(wr, 3).Value = Component.Name
    
    '���ʂƕK�v���ʏ㏑��
    Cells(wr, 4).Value = Component.Quantity * Cells(c.Row, 4).Value
    Cells(wr, 10).Value = Component.Quantity * Cells(c.Row, 4).Value
    
    '1�ڂ̃A�C�e���ɂ̂ݔ̔����i��t���ւ���
    '�����]�L�σt���O
    Dim Flg As Boolean
    
    If Component.Quantity = 1 And Flg = False Then
    
        Cells(wr, 5) = Cells(c.Row, 5).Value
        Cells(c.Row, 5).Value = 0
                
        Flg = True
        
    Else
        Cells(c.Row, 5).Value = 0
        
    End If
    
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
