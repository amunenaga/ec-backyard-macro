Attribute VB_Name = "Module1"
Option Explicit
Sub ICOKURI�A��()

    Call ConcatenateICOKURI
    
    '���sPC�̃f�X�N�g�b�v���t���p�X��
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
        
    Dim PutFolder As String
    PutFolder = CStr(wsh.SpecialFolders("desktop")) & "\ICOKURI�A���ς�\"
        
    '���s����PC�̃f�X�N�g�b�v�ɁuICOKURI�A���ς݃t�H���_�v�������ꍇ
    If Dir(PutFolder) = "" Then
    
        PutFolder = Replace(PutFolder, "ICOKURI�A���ς�\", "")
    
    End If
    
    '�}�N���N���{�^�����폜
    Sheet1.Shapes(1).Delete
    
    '�A������CSV��xlsx�`���ŕۑ�
    Application.DisplayAlerts = False
    
        ThisWorkbook.SaveAs Filename:=PutFolder & "ICOKURI" & Format(Date, "MMdd") & ".xlsx"
     
    Application.DisplayAlerts = True
    
End Sub

Private Sub ConcatenateICOKURI()
Const ICOKURI_PC As String = "\\mos10\"

Dim FolderName(2) As String
FolderName(0) = "�A�}�]����z��"
FolderName(1) = "�y�V����"
FolderName(2) = "���t�[\���t�[������"

Dim i As Integer

For i = 0 To 2

    Dim CsvPath As String
    CsvPath = FindTodaysCSV(ICOKURI_PC & FolderName(i))
    
    If CsvPath <> "" Then
        Call ImportICOKURI(CsvPath)
    End If

Next

End Sub

Sub ImportICOKURI(ByVal CsvPath)
Attribute ImportICOKURI.VB_ProcData.VB_Invoke_Func = " \n14"

'�ŏI�s��Range�A�����o����Z���̓���
Dim LastRow As Long
LastRow = Cells(1, 1).SpecialCells(xlCellTypeLastCell).Row

Dim PutStartCell As Range

If LastRow = 1 Then
    Set PutStartCell = Range("A1")
Else
    Set PutStartCell = Cells(LastRow + 1, 1)
End If

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & CsvPath, Destination:=PutStartCell _
    )
    .Name = "ICOKURI"
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
    .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
    2, 2, 2, 2, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

ActiveWorkbook.Connections(1).Delete

End Sub

Function FindTodaysCSV(ByVal CsvFolderPath As String) As String

'CSV�t�H���_�̃p�X�w��A�Ō�\�}�[�N���`�F�b�N
If Not Right(CsvFolderPath, 1) = "\" Then
    CsvFolderPath = CsvFolderPath & "\"
End If

'���s���o�C���f�B���O�Ńt�@�C���I�u�W�F�N�g
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim f As Object, TodayCSV As Object
      
'�w��t�H���_�[����FileName���܂ރt�@�C�����𒲂ׂāA�{�� ���t�t�@�C������擾����

For Each f In FSO.GetFolder(CsvFolderPath).Files

    If DateDiff("D", f.DateLastModified, DateValue(Date)) = 0 And f.Name Like "ICOKURI*" Then
    
        Set TodayCSV = f
        Exit For
        
    End If

Next

If f Is Nothing Then

      Exit Function
      
End If

FindTodaysCSV = CsvFolderPath & f.Name

End Function
