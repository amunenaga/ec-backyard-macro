Attribute VB_Name = "SetParser"
Option Explicit

Const TIED_ITEM_LIST_BOOK As String = "��ď��iؽ�.xls"
Const LIST_BOOK_FOLDER As String = "\\server02\���i��\�l�b�g�̔��֘A\"
Sub ParseItems(r As Range)

Call OpenListBook

ThisWorkbook.Activate

Dim HitSheet As Worksheet
Set HitSheet = SetParser.SearchTiedItemSheet(r.Value)

Dim ComponentItems As Collection
Set ComponentItems = GetComponentItems(r.Value, HitSheet)

'�Z�b�g���e�����o������
Call InsertComponetRow(r, ComponentItems)

End Sub

Private Sub InsertComponetRow(c As Range, d As Collection)

Dim i As Long, v As Variant

For i = 1 To d.Count
    
    Set v = d.Item(i)
    
    Rows(c.Offset(1, 0).Row).Insert (xlShiftDown)
    
    '6�P�^�����6�P�^�A�Ȃ����JAN
    c.Offset(1, 0).NumberFormatLocal = "@"
    
    If v.Code <> "" Then

        c.Offset(1, 0).Value = v.Code
    
    Else
    
        c.Offset(1, 0).Value = v.Jan
    
    End If
    
    '���i���o�͂ƕK�v���ʂ�������
    c.Offset(1, 1).Value = v.Name
    c.Offset(1, 2).Formula = "=" & v.Quantity & "*" & c.Offset(0, 2).Value
    c.Offset(1, 2).Value = c.Offset(1, 2).Value
    
    '1�ڂ̃A�C�e���ɂ̂ݔ̔����i��t���ւ���
    
    '�����]�L�σt���O
    Dim Flg As Boolean
    
    If v.Quantity = 1 And Flg = False Then
    
        c.Offset(1, 3).Value = c.Offset(0, 3).Value
        c.Offset(0, 3).Value = 0
                
        Flg = True
        
    Else
        c.Offset(1, 3).Value = 0
        
    End If
    
    '�}����̍s�ɁA���t�[�o�^�R�[�h�̓Z�b�g��7777�R�[�h������
    c.Offset(1, -1).Value = c.Value
    
    '�������A�}����̍s�ɒ����ԍ�������
    c.Offset(1, -3).Value = c.Offset(0, -3).Value
    
Next

End Sub

Private Function SearchTiedItemSheet(Code As String) As Worksheet
    '�Y���R�[�h�̂��郏�[�N�V�[�g��T���܂��B
      
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

Private Function GetComponentItems(TiedCode As String, TiedCodeList As Worksheet) As Collection

'�n���ꂽ�V�[�g�ƃR�[�h����A�Z�b�g���eCollection��Ԃ��܂��B
'�Ăяo�����ŃG���[�n���h�����O���s���̂ŁAOn Error�X�e�[�g�����g�͕s�v

'�o�^�R�[�h�̃����W�A������Match�֐��Œ��ׂāACode�̍s�ԍ����o��
Dim CodeRange As Range
Set CodeRange = TiedCodeList.Range("A1:A" & TiedCodeList.Cells(2, 1).SpecialCells(xlCellTypeLastCell).Row)

Dim HitRow As Double
HitRow = WorksheetFunction.Match(TiedCode, CodeRange, 0)


Dim ComponetItems As Collection
Set ComponetItems = New Collection

'E��=5����A�Z�b�g���e�̓X�^�[�g
'�w�b�_�[  SKU(�A��77777�n�܂�)/����(�ō�)/JAN�P�ʂ̑�����(���_���)     /JAN /����SKU /���� / ���i��

'��J�E���^
Dim i As Integer
i = TiedCodeList.Rows(1).Find("���i���1").Column

'IsEmpty���Ƌ󔒃Z���E���ꍇ������
Do Until TiedCodeList.Cells(HitRow, i) = ""

    Dim UnitCell As Range
    
    Dim Unit As ComponentItem
    Set Unit = New ComponentItem
    
    Set UnitCell = TiedCodeList.Cells(HitRow, i)
    
    With Unit
        
        .Jan = UnitCell.Value
        .Code = UnitCell.Offset(0, 1).Value
        .Name = UnitCell.Offset(0, 3).Value
        .Quantity = CLng(UnitCell.Offset(0, 2).Value)
    
    End With
        
    ComponetItems.Add Unit
    
    i = i + 4

Loop

Set GetComponentItems = ComponetItems

End Function

Private Sub OpenListBook()

'�Z�b�g���X�g�̃G�N�Z���t�@�C�����J���܂��B

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = TIED_ITEM_LIST_BOOK Then
        
        GoTo ret
    
    End If

Next

Set wb = Workbooks.Open(LIST_BOOK_FOLDER & TIED_ITEM_LIST_BOOK, ReadOnly:=True)

ret:

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

Sub ParseScalingSet(r As Variant)

Dim Code As String
Code = r.Value

Dim SeparatedCode As Variant
SeparatedCode = Split(Code, "-", 2)

'IsNumeric���\�b�h�ŁA�n�C�t���̌�낪���l�ɕϊ��\���`�F�b�N
'�ϊ��\�Ȃ�A���Z�b�g�ƌ��Ȃ�

If Not IsNumeric(SeparatedCode(1)) Then
    Exit Sub
End If

'�Z�b�g�Ȃ�AD��͒P�̃R�[�h�A���ʂ̓Z�b�g���ʁ~�󒍐���
r.NumberFormatLocal = "@"
r.Value = CStr(SeparatedCode(0))

Range("F" & r.Row).Value = Range("F" & r.Row).Value * CLng(Val(SeparatedCode(1)))

'���l�ɁA�Z�b�g�����ϋL��
Range("K" & r.Row).Value = Range("K" & r.Row).Value & "�Z�b�g���� ��"

End Sub

