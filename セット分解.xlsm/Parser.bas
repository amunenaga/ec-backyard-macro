Attribute VB_Name = "Parser"
Option Explicit

Const TIED_ITEM_LIST_BOOK As String = "���iؽ�.xls"
Const LIST_BOOK_FOLDER As String = "\\server02\���i��\�l�b�g�̔��֘A\"
Sub ParseItems()

    If Selection.Columns.Count > 1 Then
        
        MsgBox "2��ȏ�I�����Ȃ��ŉ������B"
        End
    
    End If

    Dim CurrentWorkBook As Workbook
    Set CurrentWorkBook = ActiveWorkbook
    
    Call OpenListBook

    CurrentWorkBook.Activate

    Dim rng As Range, r As Range, sh As Worksheet
    Set rng = Selection

    For Each r In rng
        
        If Not r.Value Like "#####*" Then GoTo Continue
        
        '�Z�b�g���e���擾���� �[���I��Try-Catch
        'Try
        On Error Resume Next
            
            Dim HitSheet As Worksheet
            Set HitSheet = Parser.SearchTiedItemSheet(r.Value)
        
            Dim Items As Collection
            Set Items = GetComponentItems(r.Value, HitSheet)
            
            '�Z�b�g���e�����o������
            
            Call InsertComponetRow(r, Items)
            
            'Catch
            If Err Then
                
                '�G���[�������߂����\�b�h
                
                Err.Clear
                GoTo Continue
            
            End If
            
        On Error GoTo 0
            
        '�Z�b�g�i�R�[�h�̓��X�g���ɂ��邪�A���i�o�^���Ȃ��ꍇ
        If Items.Count = 0 Then
                            
            '�G���[�������߂����\�b�h
            
            GoTo Continue
        
        End If
     
Continue:
        
    Next

End Sub

Sub InsertComponetRow(c As Range, d As Collection)

Dim v As Variant

For Each v In d
    
    Rows(c.Offset(1, 0).Row).Insert (xlShiftDown)
    
    '�Г��R�[�h����ΎГ��R�[�h�A�Ȃ����JAN
    If v.Code <> "" Then
        
        c.Offset(1, 0).Value = v.Code
    
    Else
    
        c.Offset(1, 0).Value = v.Jan
    
    End If
    
    '���i���o�͂ƕK�v���ʂ�������
    If TypeName(c.Offset(0, 1).Value) = "String" Then
        
        c.Offset(1, 1).Value = v.Name
        c.Offset(1, 2).Formula = "=" & v.Quantity & "*" & c.Offset(0, 2).Value
        
    Else
    
        c.Offset(1, 2).Value = v.Name
        c.Offset(1, 1).Formula = "=" & v.Quantity & "*" & c.Offset(0, 1).Value
    
    End If
    
    
Next

End Sub


Function SearchTiedItemSheet(Code As String) As Worksheet
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


Function GetComponentItems(TiedCode As String, TiedCodeList As Worksheet) As Collection

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
'�w�b�_�[  �A��/����(�ō�)/JAN�P�ʂ̑�����(���_���)     /JAN /�Г��R�[�h /���� / ���i��

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

Sub OpenListBook()

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

Sub CloseDataBook()

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = TIED_ITEM_LIST_BOOK Then
        
        wb.Close SaveChanges:=False
    
    End If

Next

End Sub
