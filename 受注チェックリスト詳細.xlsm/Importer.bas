Attribute VB_Name = "Importer"
Option Explicit
Sub CSV�Ǎ�()

Worksheets("Santyoku�󒍃f�[�^").Activate

Dim CsvPath As String
CsvPath = GetOrderCheckListPath()

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & CsvPath, Destination:=Range("$A$2"))
    .Name = "�󒍃`�F�b�N���X�g�ڍדǍ�"
    .FieldNames = False
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
    .TextFileColumnDataTypes = Array(2, 2, 2, 1, 9, 9, 9, 2, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, _
    9, 9, 9, 9, 5, 5, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 2, 1, 2, 9, 9, 9, 9 _
    , 9, 9, 9, 2, 9, 9, 9, 2, 2, 2, 2, 2, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, _
    9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 5, 9, 9, 9, 9, 9, 9, 9)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

ActiveWorkbook.Connections(1).Delete


'�ǂݍ��݌�A�捞���̓��t�`�F�b�N �ŏ��̒����s�ƍŌ�̒����s�̓��t�ɑ΂���

Dim LastRow As Long
LastRow = Range("Q1").SpecialCells(xlCellTypeLastCell).Row


'�{�����t�Ȃ�΁A���̏����͊���
If DateDiff("D", Cells(2, 17).Value, DateValue(Date)) = 0 _
    And DateDiff("D", Cells(LastRow, 17).Value, DateValue(Date)) = 0 Then
    
    Exit Sub

End If


'�Ǎ��f�[�^���{���捞�łȂ������ꍇ�A���s�ۂ��_�C�A���O�Ō��߂Ă��炤�B
Dim ContinueWrongDate As VbMsgBoxResult
    ContinueWrongDate = MsgBox(Buttons:=vbExclamation + vbYesNo, Prompt:="�Y���ւ̎捞�����{���ł͂���܂���B" & vbLf & "�����𑱍s���܂����H" & vbLf & vbLf & "�捞�f�[�^�L�ڂ̎捞��:" & Range("Q2").Value)

If ContinueWrongDate = vbNo Then
    
    '���s���Ȃ��ꍇ�A�f�[�^�����������̂܂܂ɂ��邩�I��
    Dim ChooseDataClear As VbMsgBoxResult
    ChooseDataClear = MsgBox(Buttons:=vbExclamation + vbYesNo, Prompt:="�Ǎ��σf�[�^���������܂����H")
    
    If ChooseDataClear = vbYes Then
        '�f�[�^�������ă}�N���S�̂��I��
        Worksheets("Santyoku�󒍃f�[�^").UsedRange.Offset(1, 0).Clear
        End
    
    Else
        '�f�[�^�m�F�̏�ő��s����ꍇ�A���s�p�{�^����ǉ��B
        With ActiveSheet.Buttons.Add(709.5, 54, 201, 42)
            .OnAction = "��ƃV�[�g�փf�[�^���o"
            .Characters.Text = "�Ǎ��σf�[�^�ŏ����𑱍s"
        End With
    
    End If
    
End If

End Sub
Private Function GetOrderCheckListPath() As String
'�{��CSV���w��t�H���_���T���B

Const SANTYOKU_DUMP_FOLDER As String = "\\Server02\���i��\�l�b�g�̔��֘A\����f�[�^\ARY�󒍃`�F�b�N���X�g\" '����\�}�[�N�K�{

'���s���o�C���f�B���O�Ńt�@�C���I�u�W�F�N�g
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim f As Object, TodayCSV As Object
      
'�w��t�H���_�[����FileName���܂ރt�@�C�����𒲂ׂāA�{�� ���t�t�@�C������擾����

For Each f In FSO.GetFolder(SANTYOKU_DUMP_FOLDER).Files

    If DateDiff("D", f.DateLastModified, DateValue(Date)) = 0 Then
    
        Set TodayCSV = f
        Exit For
        
    End If

Next

'�{�����t�t�@�C�����Ȃ���΃t�@�C���w��_�C�A���O���o���Ď蓮�Z�b�g
If TodayCSV Is Nothing Then
    
    MsgBox Prompt:="�{���̎󒍃`�F�b�N���X�g �t�@�C����������܂���ł����B" & vbLf & "�t�@�C�����w�肵�ĉ������B", _
            Buttons:=vbCritical
    
    '�J�����g�t�H���_���ړ����č���f�[�^�t�H���_�Ńt�@�C���w��_�C�A���O���J��
    '@url http://officetanaka.net/other/extra/tips15.htm
    CreateObject("WScript.Shell").CurrentDirectory = "\\server02\���i��\�l�b�g�̔��֘A\����f�[�^\"
    
    Dim FilePath As String
    FilePath = Application.GetOpenFilename()
    
    Set TodayCSV = FSO.getfile(FilePath)

End If

GetOrderCheckListPath = SANTYOKU_DUMP_FOLDER & TodayCSV.Name

End Function
