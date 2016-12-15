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


'�捞���̓��t�`�F�b�N �ŏ��̒����s�ƁA�Ō�̒����s�̓��t�ɑ΂���

Dim LastRow As Long
LastRow = Range("Q1").SpecialCells(xlCellTypeLastCell).Row

If DateDiff("D", Cells(2, 17).Value, DateValue(Date)) <> 0 _
    Or DateDiff("D", Cells(LastRow, 17).Value, DateValue(Date)) <> 0 Then

    Dim ContinueWrongDate As VbMsgBoxResult
        ContinueWrongDate = MsgBox(Buttons:=vbExclamation + vbOKCancel, Prompt:="�Y���ւ̎捞�����{���ł͂���܂���B" & vbLf & "�����𑱍s���܂��B" & vbLf & vbLf & "�捞�f�[�^�L�ڂ̎捞��:" & Range("Q2").Value)
    
    If ContinueWrongDate <> vbOK Then
        '���s���Ȃ��ꍇ�A�f�[�^�������ă}�N���I��
        Worksheets("Santyoku�󒍃f�[�^").UsedRange.Offset(1, 0).Clear
        End
        
    End If
    
End If

End Sub
Private Function GetOrderCheckListPath() As String
'�s�b�L���O�V�[�g��-a���I���̃Z�b�g����O�t�@�C����T���ăt���p�X���Z�b�g

Const SANTYOKU_DUMP_FOLDER As String = "\\Server02\���i��\�l�b�g�̔��֘A\����f�[�^\ARY�󒍃`�F�b�N���X�g\" '����\�}�[�N�K�{

'���s���o�C���f�B���O
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

'�{�����t�̃t�@�C�����Ȃ���΁A��U�}�N���I��
'TODO:�{�����t�t�@�C�����Ȃ���΃t�@�C���w��_�C�A���O���o���Ď蓮�Z�b�g
If TodayCSV Is Nothing Then End

GetOrderCheckListPath = SANTYOKU_DUMP_FOLDER & TodayCSV.Name

End Function
