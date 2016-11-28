Attribute VB_Name = "Module1"
Option Explicit

Const MAKER_QTY_PATH As String = "\\server02\���i��\�l�b�g�̔��֘A\z�݌�\���S�X���[�J�[�݌ɕ\.csv"
Const SAVE_FOLDER As String = "\\server02\���i��\�l�b�g�̔��֘A\�����֘A\��z���쐬\" '�Ō�K��\�}�[�N

'PickingSheetNames(2)�̕��тƓ����A��z�˗����i���J�E���^
Dim ItemCount(2) As Integer


Sub ���S�X��z���X�g�쐬()

'�V�[�g�ɖ{�����t������
Worksheets("���S�X�{����").Range("A1").Value = Format(Date, "m��d��")

'����Ń��S�XB2B����_�E�����[�h���Ă���݌ɕ\����荞��
Call FetchLogosQuantityCsv

'�e�s�b�L���O�V�[�g���R�s�[���āA���S�X��z�Ń}�[�N����Ă��鏤�i���R�s�[
'���t�[�̂݃s�b�L���O�V�[�g�ł͂Ȃ��AMOS10�ɕۑ������Meisai.csv���g��

'���[�����A�s�b�L���O�V�[�g�Ăяo���ƃV�[�g���Ɏg�������� �z��
Dim PickingSheetNames(2) As String
PickingSheetNames(0) = "�A�}�]��"
PickingSheetNames(1) = "�y�V"
PickingSheetNames(2) = "���t�["



Dim PickingSheetName As Variant

For Each PickingSheetName In PickingSheetNames
    
    Dim Name As String
    
    '�����q��Variant�^�iVBA�̎d�l�j�Ȃ̂�CopySheet�֐��֓n����X�g�����O�^�ɃL���X�g
    Name = CStr(PickingSheetName)
    
    Call CopySheet(Name)
    Call ExtractLogosItems(Name)
    
Next

'�Ō�ɃR�s�[�����V�[�g��Active�Ȃ̂Ŗ{�����V�[�g�ɖ߂�
Worksheets("���S�X�{����").Activate


'���S�X��z�i�̗L�����m�F
If ActiveSheet.UsedRange.Rows.Count = 1 Then

    MsgBox prompt:="���S�X �s�b�L���O�V�[�g�ł̎�z�˗����i�͂O�_�ł��B" & vbLf & "�A�b�v���[�h�p�t�@�C���͐�������܂���B"
    Exit Sub

End If

'�i�ԁA���[�J�[�݌ɂ���������Vlookup��������A�͈͂��n�[�h�R�[�f�B���O���Ă���̂Œ���
Call InsertVlookup

With ActiveSheet
    
    .UsedRange.Columns.AutoFit
    .Columns("C").ColumnWidth = 50
   
End With

'Server02�̎�z���쐬�t�H���_��xlsx�`���ŕۑ�
Application.DisplayAlerts = False
ThisWorkbook.SaveAs FileName:=SAVE_FOLDER & "���S�X" & Format(Date, "mmdd") & ".xlsx"

Call SaveAsCsv

'�{�^���������āA�������b�Z�[�W�\��
ThisWorkbook.Worksheets("���S�X�{����").Activate

ActiveSheet.Shapes("ButtonExtractLogos").Delete

MsgBox prompt:="���S�XB2B�A�b�v���[�h�t�@�C�� �ۑ�����" & vbLf & _
                "Amazon���F" & ItemCount(0) & "�_" & vbLf & _
                "�y�V���F" & ItemCount(1) & "�_" & vbLf & _
                "���t�[���F" & ItemCount(2) & "�_"

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
        GoTo Continue
    End If
    
    '�i�Ԃ����S�X�i�ԃV�[�g����E��Vlookup�������� �����W�̓n�[�h�R�[�f�C���O�����A�Z���Ɏ�������̂ł܂�������
    
    If Not IsEmpty(pc) Then GoTo Continue
         
    '6�P�^�ň�������Vlookup
    pc.Formula = "=Vlookup(" & c.Address & ",���S�X�i�ԃV�[�g!$A$1:$C$2723,3,FALSE)"
    
    If IsError(pc.Value) Then
    
        'Jan�ň�������Vlookup
        pc.Formula = "=Vlookup(" & c.Address & ",���S�X�i�ԃV�[�g!$B$1:$C$2723,2,FALSE)"
    
    End If
    
    '�i�ԃV�[�g�ł��_���Ȃ�A���S�X���[�J�[�݌ɕ\����JAN�ň�������
    If IsError(pc.Value) Then
        
        On Error Resume Next
            Dim CurRow As Double
            CurRow = WorksheetFunction.Match(pc.Value, Worksheets("���[�J�[�݌ɕ\").Range("B1:B4000"), 0)
            
            pc.Value = CStr(Worksheets("���[�J�[�݌ɕ\").Cells(CurRow, 1))
        
            If Err Then
                pc.Value = ""
                Err.Clear
            End If
        
        On Error GoTo 0
    
    End If
    
    '�Z�b�g���e�̏��i�͏��i������s�ɂȂ�̂ŁAVlookup�ň�������
    If Not TypeName(c.Offset(0, 1).Value) = "String" Then
        c.Offset(0, 1).Formula = "=Vlookup(" & pc.Address & ",���S�X�i�ԃV�[�g!$C$1:$D$2723,2,FALSE)"
    End If
    
    '���S�X ���[�J�[�݌ɐ�����������
    pc.Offset(0, 1).Formula = "=Vlookup(" & pc.Address & ",���[�J�[�݌ɕ\!A:E,4,FALSE)"
    
Continue:
    i = i + 1

Loop

End Sub

Private Sub ExtractLogosItems(PickingSheetName As String)
'���S�X���i�̖{����z�V�[�g�ւ̒��o

Dim TodayDate As String
TodayDate = Format(Date, "mmdd")

Worksheets(PickingSheetName & TodayDate).Activate

'�s�b�L���O�V�[�g�ʂ̏���
'�y�V�A���̂������E�R���N�g�ɉ��F�̔w�i�F�����Ă���̂ŁA�F������
If PickingSheetName = "�y�V" Then

    Dim AnnotationHeader As Range
    Set AnnotationHeader = Range("A1:E20").Find("�����E�R���N�g*")
    
    If Not AnnotationHeader Is Nothing Then
    
        AnnotationHeader.Interior.ColorIndex = 0
        
    End If

End If

'���i���̗�A�s�ԍ������
Dim FoundCell As Range
Set FoundCell = Range("A1:E20").Find("���i��")

Dim col As Double, nrow As Double
col = FoundCell.Column
nrow = FoundCell.Row

'�t�B���^�[���郌���W���w��
Dim ProductListRange As Range
Set ProductListRange = Range(Cells(2, 1), Range("A1").CurrentRegion.SpecialCells(xlCellTypeLastCell))

'�F�Ńt�B���^�[�A���t�[�̂݃��S�X������ł̃t�B���^�[
'�s�b�L���O�V�[�g�ł̓��S�X�̎�z�˗��͔w�i�F�����F


If PickingSheetName = "���t�[" Then

    ProductListRange.AutoFilter Field:=5, Criteria1:="���S�X*"
    
Else
    
    ProductListRange.AutoFilter Field:=col, Criteria1:=RGB(255, 255, 0), Operator:=xlFilterCellColor

End If

'�t�B���^�[������̍s�����J�E���g���˗����i��
Dim CountItem As Long
CountItem = WorksheetFunction.Subtotal(3, Cells(3, col).Resize(Cells(3, col).SpecialCells(xlCellTypeLastCell).Row, 1))

Call setItemCount(PickingSheetName, CountItem)

'�t�B���^�[���ĕ\�����Ă��郌���W�̂ݎ擾
Dim A As Range, B As Range
Set A = ProductListRange.SpecialCells(xlCellTypeVisible)

'���i���̑O��1�񁁌v3����R�s�[�������A1��O���R�[�h�A1���끁����
Set B = Cells(nrow, col).Offset(1, -1).Resize(Cells(2, col).SpecialCells(xlCellTypeLastCell).Row, 3)

Dim IntersectRange As Range
Set IntersectRange = Application.Intersect(A, B)

'�t�B���^�[���ĕ\�����Ă��镔�������S�X�{�����փR�s�[
Dim LastRow As Integer
LastRow = Worksheets("���S�X�{����").Range("A1").SpecialCells(xlCellTypeLastCell).Row

Dim CopyDestinationRange As Range
Set CopyDestinationRange = Worksheets("���S�X�{����").Cells(LastRow + 1, 1).Offset(0, 1)

'�󒍖������ƃt�B���^�[��̃����W���Ȃ��̂Ń`�F�b�N
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

'�A�}�]���̃s�b�L���O�V�[�g�t�@�C�����́u�s�b�L���O�v
If Mall = "�A�}�]��" Then Mall = "�s�b�L���O"

'���t�[�̂�MOS10�̎�CSV��ǂ݂ɍs��
If Mall = "���t�[" Then
    Call FetchYahooMeisai
    Exit Sub
End If

Workbooks.Open FileName:=RetrievePickingFilePath(Mall), ReadOnly:=True

Dim BookName As String
BookName = ActiveWorkbook.Name '�t�@�C�����J������J�����u�b�N��Active�ɂȂ��Ă���

ActiveWorkbook.Sheets(1).Copy After:=ThisWorkbook.Worksheets("���S�X�{����")

If Mall = "�s�b�L���O" Then Mall = "�A�}�]��" '�V�[�g���̓A�}�]��

ActiveSheet.Name = Mall & Format(Date, "mmdd")

Workbooks(BookName).Close SaveChanges:=False

End Sub

Private Function RetrievePickingFilePath(FileName As String) As String
'�s�b�L���O�V�[�g��-a���I���̃Z�b�g����O�t�@�C����T���ăt���p�X���Z�b�g

Const PICKING_FILE_FOLDER As String = "\\Server02\���i��\�l�b�g�̔��֘A\�s�b�L���O\" '����\�}�[�N�K�{

'�y�V�̏ꍇ�A�y�VP�V�[�g0627-a.xls

'���s���o�C���f�B���O ScriptingRuntime��Dictionary�z��g���̂ɕK�v�ŎQ��ON������A���O�o�C���f�B���O�ł��������B
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
    
Dim f As Object, Newest As Object
      
'���O�o�C���f�B���O
'Dim FSO As FileSystemObject
'Set FSO = New FileSystemObject

'Dim f As File, Newest As File


'�w��t�H���_�[����FileName���܂ރt�@�C�����𒲂ׂāA�ŐV�̃t�@�C����1�擾����B
'LINQ�������A1�\���ōςނ̗~����

For Each f In FSO.GetFolder(PICKING_FILE_FOLDER).Files

    If f.Name Like FileName & "*-a.xls*" Then
    
        Set Newest = f
    
        Exit For
    End If

Next


RetrievePickingFilePath = PICKING_FILE_FOLDER & Newest.Name

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


Private Sub FetchLogosQuantityCsv()

Worksheets("���[�J�[�݌ɕ\").Activate '.QueryTables���\�b�h��ActiveSheet�łȂ��Ƒ���Ȃ�

With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & MAKER_QTY_PATH, Destination:=Range("$A$1"))
        
        .Name = "���S�X���[�J�[�݌ɕ\"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 932
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(2, 2, 2, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False

End With

Range("G1").Value = "z�݌ɂ�CSV�擾����"
Range("G2").Value = Hour(Time) & ":" & Minute(Time)

Worksheets("���S�X�{����").Activate

End Sub

Private Sub FetchYahooMeisai()
'���t�[�̂�MOS10��Meisai.CSV��ǂݍ���ŃV�[�g���C������B

'�V�[�g�}���ʒu�̓��S�X�{���̌��
ThisWorkbook.Worksheets.Add After:=ThisWorkbook.Worksheets("���S�X�{����")
ActiveSheet.Name = "���t�[" & Format(Date, "mmdd")

    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;\\MOS10\Users\mos10\Desktop\���t�[\Meisai.csv", Destination:=Range("$A$1") _
        )
        
        .Name = "Meisai"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 932
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = ","
        .TextFileColumnDataTypes = Array(2, 1, 1, 2, 2, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    
    End With
    
    'ExtractLogosItems���\�b�h�Œ��o�ł���悤�A�V�[�g�𒲐߂��܂��B
    
    ActiveSheet.Rows(1).Find("Description").Value = "���i��"
    Columns("C:C").Copy
    Columns("F:F").Insert Shift:=xlRight
    Rows(1).Insert

End Sub

Private Function getMallId(MallName As String) As String

Dim MallId As String

Select Case MallName
    
    Case "�A�}�]��"
        MallId = "A"
    
    Case "�y�V"
        MallId = "R"
    
    Case "���t�["
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

Private Sub SaveAsCsv()

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

Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs FileName:=SAVE_FOLDER & "���S�X�����o�^CSV" & Format(Date, "mmdd"), FileFormat:=xlCSV

Application.DisplayAlerts = True

End Sub

Private Sub setItemCount(ByVal MallName As String, ByVal Count As Long)

Select Case MallName
    
    Case "�A�}�]��"
        ItemCount(0) = Count
    
    Case "�y�V"
        ItemCount(1) = Count
    
    Case "���t�["
        ItemCount(2) = Count
    
    End Select

End Sub
