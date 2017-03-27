Attribute VB_Name = "Module1"
Option Explicit

Const MAKER_QTY_PATH As String = "\\server02\���i��\�l�b�g�̔��֘A\z�݌�\���S�X���[�J�[�݌ɕ\.csv"
Const SAVE_FOLDER As String = "\\server02\���i��\�l�b�g�̔��֘A\�����֘A\��z���쐬\" '�Ō�K��\�}�[�N

'PickingSheetNames(2)�̕��тƓ����A��z�˗����i���J�E���^
Dim ItemCount(2) As Integer

Sub ���S�X��z���X�g�쐬()

'�}�N���N���{�^���폜�A�Ē��o�{�^���폜
ActiveSheet.Shapes("ButtonExtractLogos").Delete

'�V�[�g�ɖ{�����t������
Worksheets("���S�X�{����").Range("A1").Value = Format(Date, "m��d��")

'�e�s�b�L���O�V�[�g���R�s�[���āA���S�X��z�Ń}�[�N����Ă��鏤�i���R�s�[

'���[�����A�s�b�L���O�V�[�g�Ăяo���ƃV�[�g���Ɏg�������� �z��
Dim PickingSheetNames(2) As String
PickingSheetNames(0) = "Amazon"
PickingSheetNames(1) = "�y�V"
PickingSheetNames(2) = "Yahoo"

Dim PickingSheetName As Variant

For Each PickingSheetName In PickingSheetNames
    
    Dim Name As String
    
    '�����q��Variant�^�iVBA�̎d�l�j�Ȃ̂�CopySheet�֐��֓n����X�g�����O�^�ɃL���X�g
    Name = CStr(PickingSheetName)
    
    On Error Resume Next
        
        Call CopySheet(Name)
        Call ExtractLogosItems(Name)
    
        If Err Then
        
            MsgBox Prompt:=Name & " �I�����s�b�L���O�V�[�g�����A�����𑱍s���܂��B"
        
        End If
    
    On Error GoTo 0
continue:

Next

'�Ō�ɃR�s�[�����V�[�g��Active�Ȃ̂Ŗ{�����V�[�g�ɖ߂�
Worksheets("���S�X�{����").Activate


'���S�X��z�i�̗L�����m�F
If ActiveSheet.UsedRange.Rows.Count = 1 Then

    MsgBox Prompt:="���S�X �s�b�L���O�V�[�g�ł̎�z�˗����i�͂O�_�ł��B" & vbLf & "�A�b�v���[�h�p�t�@�C���͐�������܂���B"
    Exit Sub

Else
    MsgBox Prompt:="���S�X��z�˗� ���o����" & vbLf & _
                "Amazon���F" & ItemCount(0) & "�_" & vbLf & _
                "�y�V���F" & ItemCount(1) & "�_" & vbLf & _
                "���t�[���F" & ItemCount(2) & "�_"

End If

'�i�ԁA���[�J�[�݌ɂ���������Vlookup��������
Call InsertVlookup

With ActiveSheet
    
    .UsedRange.Columns.AutoFit
    .Columns("C").ColumnWidth = 50
   
End With

'�Ē��o�{�^���ACSV�ĕۑ��{�^����z�u
'CSV�ĕۑ��{�^���z�u
With ActiveSheet.Buttons.Add( _
        Range("H3").Left, _
        Range("H3").Top, _
        Range("H3:I4").Width, _
        Range("H3:I4").Height _
        )
    .OnAction = "B2B�pCSV�ۑ�"
    .Characters.Text = "B2B�pCSV �ĕۑ�"
End With

'Server02�̎�z���쐬�t�H���_��xlsx�`���ŕۑ�
Application.DisplayAlerts = False
ThisWorkbook.SaveAs FileName:=SAVE_FOLDER & "���S�X" & Format(Date, "mmdd") & ".xlsx"

ThisWorkbook.Worksheets("���S�X�{����").Activate

'�i�Ԃ��擾�ł��Ă��Ȃ����i�������CSV��������U�ۗ��A�{�^���Ŏ蓮�����Ƃ���
If HasProductCode Then

    Call B2B�pCSV�ۑ�
    MsgBox Prompt:="B2B�pCSV��ۑ����܂����B", Buttons:=vbInformation

Else
    
    MsgBox Prompt:="���S�X�i�Ԃ��擾�ł��Ă��Ȃ����i������܂��B" & vbLf & "�m�F��A�ĕۑ��{�^����CSV�𐶐����ĉ������B", Buttons:=vbCritical

End If

End Sub

Private Sub InsertVlookup()
'�i�Ԃ����S�X�i�ԃV�[�g������������Ă���Vlookup��������

Worksheets("���S�X�{����").Activate

Dim i As Integer

i = 2

Do Until IsEmpty(Cells(i, 2))
        
    Dim c As Range, pc As Range
    Set c = Cells(i, 2) '�R�[�h�Z��
    Set pc = Cells(i, 5) '�i�ԃZ��
    
    c.NumberFormatLocal = "@"
    c.Value = CStr(c.Value)
      
    '�Z�b�g�R�[�h�𕪉�����
    If c.Value Like "77777*" Then
        Call MarkAsTiedItem(c)
        Call InsertComponentItems(c)
        GoTo continue
    End If
    
    '�i�Ԃ����S�X�i�ԃV�[�g����E��Vlookup��������
    
    Dim LastRow As Long
    LastRow = Worksheets("���S�X�i�ԃV�[�g").UsedRange.Rows.Count
    
    If Not IsEmpty(pc) Then GoTo continue
         
    '6�P�^�ň�������Vlookup
    pc.Formula = "=Vlookup(" & c.Address & ",���S�X�i�ԃV�[�g!$A$1:$C$" & LastRow & ",3,FALSE)"
    
    If IsError(pc.Value) Then
    
        'Jan�ň�������Vlookup
        pc.Formula = "=Vlookup(" & c.Address & ",���S�X�i�ԃV�[�g!$B$1:$C$" & LastRow & ",2,FALSE)"
    
    End If
    
    '�i�ԃV�[�g�ł��_���Ȃ�A���S�X���[�J�[�݌ɕ\����JAN�ň�������
    If IsError(pc.Value) Then
        
        On Error Resume Next
            Dim CurRow As Double
            CurRow = WorksheetFunction.Match(pc.Value, Worksheets("���[�J�[�݌ɕ\").Range("B1:B8000"), 0)
            
            pc.Value = CStr(Worksheets("���[�J�[�݌ɕ\").Cells(CurRow, 1))
        
            If Err Then
                pc.Value = ""
                Err.Clear
            End If
        
        On Error GoTo 0
    
    End If
    
    '�Z�b�g���e�̏��i�͏��i������s�ɂȂ�̂ŁAVlookup�ň�������
    If Not TypeName(c.Offset(0, 1).Value) = "String" Then
        c.Offset(0, 1).Formula = "=Vlookup(" & pc.Address & ",���S�X�i�ԃV�[�g!$C$1:$D$" & LastRow & ",2,FALSE)"
    End If
    
    '���S�X ���[�J�[�݌ɐ�����������
    pc.Offset(0, 1).Formula = "=Vlookup(" & pc.Address & ",���[�J�[�݌ɕ\!A:E,4,FALSE)"
    
continue:
    i = i + 1

Loop

End Sub

Private Sub ExtractLogosItems(PickingSheetName As String)
'���S�X���i�̖{����z�V�[�g�ւ̒��o

Dim TodayDate As String
TodayDate = Format(Date, "mmdd")

Worksheets(PickingSheetName & TodayDate).Activate

'���i���̗�A�s�ԍ������
Dim FoundCell As Range
Set FoundCell = Range("A1:E20").Find("���i��")

Dim col As Double, nrow As Double
col = FoundCell.Column
nrow = FoundCell.Row

'�t�B���^�[���郌���W���w�肵�Ĕw�i�F���F���t�B���^�[
Dim ProductListRange As Range
Set ProductListRange = Range(Cells(1, 1), Range("A1").CurrentRegion.SpecialCells(xlCellTypeLastCell))
ProductListRange.AutoFilter Field:=col, Criteria1:=RGB(255, 255, 0), Operator:=xlFilterCellColor

'�t�B���^�[������̍s�����J�E���g���˗����i��
Dim CountItem As Long
CountItem = WorksheetFunction.Subtotal(2, Cells(3, col).Resize(Cells(2, col).SpecialCells(xlCellTypeLastCell).Row, 1))

Call setItemCount(PickingSheetName, CountItem)

'�t�B���^�[���ĕ\�����Ă��郌���W�̂ݎ擾
Dim A As Range, b As Range
Set A = ProductListRange.SpecialCells(xlCellTypeVisible)

'���i���̑O��1�񁁌v3����R�s�[�������A1��O���R�[�h�A1���끁����
Set b = Cells(nrow, col).Offset(1, -1).Resize(Cells(2, col).SpecialCells(xlCellTypeLastCell).Row, 3)

Dim IntersectRange As Range
Set IntersectRange = Application.Intersect(A, b)

'�t�B���^�[���ĕ\�����Ă��镔�������S�X�{�����փR�s�[
Dim LastRow As Integer
LastRow = Worksheets("���S�X�{����").Range("A1").SpecialCells(xlCellTypeLastCell).Row

Dim CopyDestinationRange As Range
Set CopyDestinationRange = Worksheets("���S�X�{����").Cells(LastRow + 1, 1).Offset(0, 1)

'�󒍖������ƃt�B���^�[��̃����W���Ȃ��̂Ń`�F�b�N���ăR�s�[
If Not IntersectRange Is Nothing Then

    CopyDestinationRange.Offset(0, -1).Value = getMallId(PickingSheetName)
    
    IntersectRange.Copy
    CopyDestinationRange.PasteSpecial Paste:=xlPasteValues
    
End If

ActiveSheet.Range("A2").CurrentRegion.AutoFilter

Exit Sub

NoMatch:

    MsgBox PickingSheetName & "�̃s�b�L���O�V�[�g�̌��o���u���i���v��������܂���ł����B"

End Sub

Private Sub CopySheet(Mall As String)

Workbooks.Open FileName:=GetPickingFilePath(Mall), ReadOnly:=True

Dim BookName As String
BookName = ActiveWorkbook.Name '�t�@�C�����J������J�����u�b�N��Active�ɂȂ��Ă���

ActiveWorkbook.Sheets(1).Copy After:=ThisWorkbook.Worksheets("���S�X�{����")

ActiveSheet.Name = Mall & Format(Date, "mmdd")

Workbooks(BookName).Close SaveChanges:=False

End Sub

Private Function GetPickingFilePath(MallName As String) As String
'�s�b�L���O�V�[�g��-a���I���̃t�@�C����T���ăt���p�X���Z�b�g

Dim FileName As String
If MallName = "Amazon" Then
    FileName = "�s�b�L���O"
ElseIf MallName = "�y�V" Then
    FileName = "�y�V"
ElseIf MallName = "Yahoo" Then
    FileName = "���t�["
End If

Const PICKING_FILE_FOLDER As String = "\\Server02\���i��\�l�b�g�̔��֘A\�s�b�L���O\" '����\�}�[�N�K�{

'���s���o�C���f�B���O ScriptingRuntime��Dictionary�z��g���̂ɕK�v�ŎQ��ON������A���O�o�C���f�B���O�ł��������B
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
    
Dim f As Object, Target As Object

'�w��t�H���_�[����FileName���܂ރt�@�C�����𒲂ׂāA���[������-a���܂�Excel�t�@�C������擾����B
'�y�V�̏ꍇ�A�y�VP�V�[�g0627-a.xls���擾

For Each f In FSO.GetFolder(PICKING_FILE_FOLDER).Files

    If f.Name Like FileName & "*-a.xls*" Then
    
        Set Target = f
    
        Exit For
    End If

Next

GetPickingFilePath = PICKING_FILE_FOLDER & Target.Name

End Function

Private Sub InsertComponentItems(c As Range)
'7777�n�܂��Cell��n���Ă�����āACode���p�[�X���ĕԂ��Ă���Dictionary�ɑ΂��āA�s��}������6�P�^�E���ʂ��o��
'�vScriptingRuntime�Q�ƁADictionary�z��̎��O�o�C���f�B���O�ɕK�{

Dim Items As Dictionary
Set Items = ParseTiedItems(c.Value)

Dim OrderedQty As Integer
OrderedQty = c.Offset(0, 2)

Dim v As Variant

For Each v In Items
    
    Rows(c.Offset(1, 0).Row).Insert (xlShiftDown)
    
    c.Offset(1, 0).Value = v
    c.Offset(1, 2).Value = Items(v) * OrderedQty

Next

End Sub

Private Function ParseTiedItems(SetCode As String) As Dictionary

'TiedItems�ŃZ�b�g���i���e�c�ł����̂��ȁB���t�����肠�킹��K�v������B
'SetItems���ƁAGetSetItems�Ƃ��ɂȂ萦������킵�� get/set���\�b�h�Ɣ��̂ŕ���킵��
'TiedItems�́A����Ō��������������i�݂����ȃC���[�W�ɂȂ��Ă��܂���

Dim TiedCodeList As Worksheet
Set TiedCodeList = Worksheets("���S�X�Z�b�g���i���X�g")

'�o�^�R�[�h�̃����W�A������Match�֐��Œ��ׂāACode�̍s�ԍ����o��
Dim CodeRange As Range
Set CodeRange = TiedCodeList.Range("A1:A" & TiedCodeList.Cells(2, 1).SpecialCells(xlCellTypeLastCell).Row)

On Error Resume Next

    Dim HitRow As Double
    HitRow = WorksheetFunction.Match(SetCode, CodeRange, 0)

On Error GoTo 0

'�R�[�h�Ńq�b�g�����s���AF��->J��->M��E�E�E�ƒ��ׂāA�R�[�h�ƌ���Dictionary�z��Ɋi�[����
Dim d As Dictionary
Set d = New Dictionary

'F��=6����A�Z�b�g���e�̓X�^�[�g
Dim i As Integer
i = 6

Do Until TiedCodeList.Cells(HitRow, i) = "" 'IsEmpty���Ƌ󔒃Z���E���ꍇ������

    Dim CodeCell As Range, Code As String, Qty As Integer
    
    Set CodeCell = TiedCodeList.Cells(HitRow, i)
    
    Code = CodeCell.Value
    Qty = CInt(CodeCell.Offset(0, 1))
    
    d.Add Code, Qty   '�Z�b�g���e���i�Ń_�u�肪����ƃG���[�Ŏ~�܂�B
    
    i = i + 4 '�V�[�g�ł�4��ŃZ�b�g���e1���i

Loop

Set ParseTiedItems = d

End Function

Private Function getMallId(MallName As String) As String

Dim MallId As String

Select Case MallName
    
    Case "Amazon"
        MallId = "A"
    
    Case "�y�V"
        MallId = "R"
    
    Case "Yahoo"
        MallId = "Y"
       
    Case Else
        MallId = "S"

End Select

getMallId = MallId

End Function

Private Sub MarkAsTiedItem(c As Range)

With c.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent6
    .TintAndShade = 0.599993896298105
    .PatternTintAndShade = 0
End With

End Sub

Private Function HasProductCode() As Boolean

Worksheets("���S�X�{����").Activate

Dim i As Long, tmp As Boolean
i = 2

'��U�Atmp�t���O��True�ŃZ�b�g�AE�񁁕i�ԗ�ɋ󗓂��G���[������ꍇ�̂�False�Ƃ���B
tmp = True

Do
    If IsError(Cells(i, 5).Value) Then
        
        tmp = False
        Exit Do
        
    ElseIf Cells(i, 5).Value = "" Then
        
        tmp = False
        Exit Do
    
    End If
    
    i = i + 1

Loop Until Cells(i, 2).Value = ""

HasProductCode = tmp

End Function

Sub B2B�pCSV�ۑ�()

'CSV�t�@�C�����J���Ă���Ε���
Application.DisplayAlerts = False

Dim wb As Workbook
For Each wb In Workbooks
    If InStr(wb.Name, "���S�X�����o�^CSV") > 0 Then wb.Close
Next

Workbooks.Add

'CSV�V�[�g�̍s�J�E���^
Dim i As Long
i = 1

With ThisWorkbook.Worksheets("���S�X�{����")
    
    Dim LastRow As Long
    LastRow = .UsedRange.Rows.Count
    
    '�i��/���ʂ𗬂�����
    Dim k As Long
    
    For k = 1 To LastRow - 1
        
        If .Range("E1").Offset(k, 0).Value <> "" Then
        
            Cells(i, 1).Value = .Range("E1").Offset(k, 0).Value
            Cells(i, 2).Value = .Range("E1").Offset(k, -1).Value
            
            i = i + 1
        
        End If
        
    Next

End With


    
ActiveWorkbook.SaveAs FileName:=SAVE_FOLDER & "���S�X�����o�^CSV" & Format(Date, "mmdd"), FileFormat:=xlCSV

Application.DisplayAlerts = True

End Sub

Private Sub setItemCount(ByVal MallName As String, ByVal Count As Long)

Select Case MallName
    
    Case "Amazon"
        ItemCount(0) = Count
    
    Case "�y�V"
        ItemCount(1) = Count
    
    Case "���t�["
        ItemCount(2) = Count
    
    End Select

End Sub
