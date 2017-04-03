Attribute VB_Name = "LoadOrderCsv"
Option Explicit

Sub LoadCsv(Optional ByVal bool As Boolean)
'�N���X���[������_�E�����[�h����CSV�Ǎ�

'�N���X���[����CSV���_�E�����[�h���Ă���t�H���_�ֈړ��A�t�@�C���w��_�C�A���O���J��
'@url http://officetanaka.net/other/extra/tips15.htm
CreateObject("WScript.Shell").CurrentDirectory = "\\server02\���i��\�l�b�g�̔��֘A\�s�b�L���O\�N���X���[��\"

Dim FilePath As String
FilePath = Application.GetOpenFilename("�N���X���[��CSV,*.csv", 2, "�N���X���[���̃s�b�L���OCSV���w��")

If FilePath = "False" Then
    MsgBox "�t�@�C���w�肪�L�����Z������܂����B" & vbLf & "�}�N�����I�����܂��B"
    End
End If

If DateDiff("D", FileDateTime(FilePath), Date) <> 0 Then
    Dim IsContinue As Integer
    IsContinue = MsgBox(prompt:="�{���̃_�E�����[�h�t�@�C���ł͂���܂���B" & vbLf & "�s�b�L���O�V�[�g�𐶐����܂����H", Buttons:=vbYesNo + vbQuestion)

    If IsContinue = vbNo Then
        MsgBox "�������I�����܂��B"
        End
    End If
End If

'�}�N���N���{�^���폜
OrderSheet.Shapes(1).Delete

'�f�[�^�ڑ��𗘗p����CSV�f�[�^��ǂݍ���
With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & FilePath, Destination:=Range("$A$2"))
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
    
    .TextFileColumnDataTypes = Array(2, 2, 2, 2, 1, 9, 9, 9, 9, 1, 1, 9, 9, 9, 1)
    
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=True
End With

ActiveWorkbook.Connections(1).Delete

'�N���X���[����CSV���ǂݍ��܂ꂽ���`�F�b�N �N���X���[�����ō̔Ԃ���A�Ԃ͐���8�P�^
If Not Range("A2").Value Like String(8, "#") Then
    MsgBox prompt:="�Ǎ��񂾃t�@�C���ɃN���X���[���̘A�Ԃ�����܂���B" & vbLf & "�������I�����܂��B", Buttons:=vbCritical, Title:="�������Ȃ��t�@�C��"
    End
End If

End Sub
