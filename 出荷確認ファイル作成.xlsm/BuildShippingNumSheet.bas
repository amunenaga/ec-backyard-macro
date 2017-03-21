Attribute VB_Name = "BuildShippingNumSheet"
Option Explicit
Sub ���샄�}�g_�V�[�g�쐬()

'�{�^���폜
Worksheets("�g�b�v").Shapes(1).Delete

'TSV/CSV�t�@�C���p�X�w��
Dim Path As Collection
Set Path = GetCsvPath()

'�f�[�^�ǂݍ���
Call LoadAmazon(Path.Item("Amazon"))
Call LoadRakuten(Path.Item("Rakuten"))
Call LoadYahoo(Path.Item("Yahoo"))

'�^����ЕʂɃV�[�g�փR�s�[
Call SortByCarrier("����}��")
Call SortByCarrier("���}�g�^�A")

'�񕝒���
Dim i As Long
For i = 1 To Worksheets.Count
    Worksheets(i).Range("A1").CurrentRegion.Columns.AutoFit
Next i

'�㏈���A�f�[�^�����N�폜�A�Z���́u���O�v�폜
Dim qt As QueryTable
For Each qt In Worksheets("�g�b�v").QueryTables
    qt.Delete
Next qt

Dim nm As Name
For Each nm In ActiveWorkbook.Names
    nm.Delete
Next nm

'�t�@�C���ۑ�
Application.DisplayAlerts = False
    ThisWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\�o�׊m�F_" & Format(Date, "yyyyMMdd") & ".xlsx", FileFormat:=xlWorkbookDefault
Application.DisplayAlerts = True

End Sub

Sub LoadAmazon(ByVal Path As String)

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & Path, Destination:=Range("$B$2")) '�p�X�͓��I�ɁA�����o�����B2�Œ�BAmazon�����荞�ނ̂�
    .Name = "Amazon"
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
    .TextFileStartRow = 4
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = True
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = False
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(2, 9, 9, 9, 9, 9, 2, 9, 9)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

Call FillMallName("Amazon")

End Sub

Sub LoadRakuten(ByVal Path As String)

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & Path, Destination:=Range("B1").End(xlDown).Offset(1, 0)) '�p�X�Ə����o����͓��I�Ɍ��߂�
    .Name = "�y�V"
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
    .TextFileStartRow = 2
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = False
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = True
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(2, 9, 9, 2, 9)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

Call FillMallName("�y�V")

End Sub

Sub LoadYahoo(ByVal Path As String)

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & Path, Destination:=Range("B1").End(xlDown).Offset(1, 0)) '�p�X�Ə����o����͓��I�Ɍ��߂�
    .Name = "yahoo"
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
    .TextFileStartRow = 2
    .TextFileParseType = xlDelimited
    .TextFileTextQualifier = xlTextQualifierDoubleQuote
    .TextFileConsecutiveDelimiter = False
    .TextFileTabDelimiter = False
    .TextFileSemicolonDelimiter = False
    .TextFileCommaDelimiter = True
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(2, 9, 2, 9, 9, 9)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

Call FillMallName("Yahoo")

End Sub

Private Sub FillMallName(ByVal MallName As String)
'CSV�ǂݍ��݌��A������[�����Ŗ��߂܂��B

Dim StartRow As Double, EndRow As Double, i As Double
StartRow = IIf(Range("A2").Value = "", 2, Range("A1").End(xlDown).Row + 1)
EndRow = Range("B1").End(xlDown).Row

For i = StartRow To EndRow
    Cells(i, 1).Value = MallName
Next i

End Sub

Sub SortByCarrier(ByVal CarrierName As String)
'�^����Ж����󂯎���āA�^����Ж��̃V�[�g�֑����ԍ����R�s�[

'�^����Ђƃt�B���^�[�����̃}�b�s���O
Dim Criteria As Variant

Select Case CarrierName
    
    Case "����}��"
        Criteria = "4031*"
    
    Case "���}�g�^�A"
        Criteria = Array("7645*", "3046*")

End Select

'�����ԍ����t�B���^�[���ăR�s�[
With Range("A1").CurrentRegion
    .AutoFilter Field:=3, Criteria1:=Criteria, Operator:=xlFilterValues
    .Copy Worksheets(CarrierName).Range("A1")
    .AutoFilter '�I�[�g�t�B���^�[����
End With

End Sub

Function GetCsvPath() As Collection
'�����ԍ�CSV�p�X���擾�A3�t�@�C���܂œ����w��\

Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd
    .Filters.Clear
    .Filters.Add "Amazon,�y�V,Yahoo!", "*.tsv; *.csv"
    .InitialFileName = "\\Server02\���i��\�l�b�g�̔��֘A\�o�גʒm"

    .Show
    
    If .SelectedItems.Count >= 4 Then
       MsgBox "�t�@�C���w�肪3�𒴂��Ă��܂��B"
       End
    End If
    
    Dim Paths As Collection, CurrentPath As String, i As Long
    Set Paths = New Collection
    
    For i = 1 To 3
        CurrentPath = fd.SelectedItems.Item(i)
        Select Case True
            Case CurrentPath Like "*amazon*"
                Paths.Add Item:=CurrentPath, Key:="Amazon"
            
            Case CurrentPath Like "*�y�V*"
                Paths.Add Item:=CurrentPath, Key:="Rakuten"
                
            Case CurrentPath Like "*yahoo*"
                Paths.Add Item:=CurrentPath, Key:="Yahoo"
            
        End Select
    Next

End With

Set GetCsvPath = Paths

End Function
