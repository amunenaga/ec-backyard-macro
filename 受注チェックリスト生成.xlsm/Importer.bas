Attribute VB_Name = "Importer"
Option Explicit
Sub CSV�Ǎ�()

Worksheets("�󒍃f�[�^").Activate

Dim CsvPath As String
CsvPath = GetOrderCheckListPath()

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & CsvPath, Destination:=Range("$A$2"))
    .Name = "�󒍃`�F�b�N���X�g�ڍדǍ�"
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
    .TextFileColumnDataTypes = Array(2, 2, 2, 1, 1, 5, 5, 2, 2, 2, 2, 2, 2, 2, 2)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=True
End With

ActiveWorkbook.Connections(1).Delete

'�}�N���N���{�^���폜
Worksheets("�󒍃f�[�^").Shapes(1).Delete

End Sub
Private Function GetOrderCheckListPath() As String
'�t�H���_���w�肵�ăt�@�C���w��_�C�A���O����t�@�C���w��

Const CSV_DL_FOLDER As String = "\\server02\���i��\�l�b�g�̔��֘A\�s�b�L���O\�N���X���[���e�X�g" '����\�}�[�N�K�{

Dim FilePath As String

'�J�����g�t�H���_���ړ����č���f�[�^�t�H���_�Ńt�@�C���w��_�C�A���O���J��
'@url http://officetanaka.net/other/extra/tips15.htm
CreateObject("WScript.Shell").CurrentDirectory = CSV_DL_FOLDER

FilePath = Application.GetOpenFilename("�N���X���[��CSV,*.csv", 2, "�N���X���[���̃s�b�L���OCSV���w��")

If FilePath = "False" Then
    MsgBox "�t�@�C���w�肪�L�����Z������܂����B" & vbLf & "�}�N�����I�����܂��B"
    End
End If

GetOrderCheckListPath = FilePath

End Function
