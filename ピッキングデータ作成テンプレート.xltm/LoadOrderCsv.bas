Attribute VB_Name = "LoadOrderCsv"
Option Explicit

Sub LoadCsv(Optional ByVal bool As Boolean)

Worksheets("�󒍃f�[�^�V�[�g").Activate

Dim FilePath As String
FilePath = Application.GetOpenFilename(Title:="CSV���w��", FileFilter:="�N���X���[��CSV,*.csv", FilterIndex:="2")

If FilePath = "False" Then
    MsgBox "�t�@�C���w�肪�L�����Z������܂����B" & vbLf & "�}�N�����I�����܂��B"
    End
End If

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

'�A�h�C���p�̃R�[�h�C���A�Z�b�g����
Call FixForAddin
Call SetParse

ActiveWorkbook.Connections(1).Delete

End Sub

Private Sub FixForAddin()
Dim CodeRange As Range, c As Range
Set CodeRange = Range(Cells(2, 2), Cells(Range("B1").SpecialCells(xlCellTypeLastCell).Row, 2))

'�A�h�C���p�̃R�[�h���L������
For Each c In CodeRange
    
    Dim CurrentCodeCell As Range
    Set CurrentCodeCell = c
    
    'I��A�A�h�C�����s�p��6�P�^�������R�[�h�A��������JAN������
    Dim ForAddinCell As Range
    Set ForAddinCell = Cells(c.Row, 9)
    
    ForAddinCell.NumberFormatLocal = "@"
    
    '6�P�^�Ȃ炻�̂܂ܓ����
    If CurrentCodeCell.Value Like String(6, "#") Then
        ForAddinCell.Value = CurrentCodeCell.Value
    
    '����5�P�^�͓��Ƀ[����ǋL
    ElseIf CurrentCodeCell.Value Like String(5, "#") Then
        
        ForAddinCell.Value = "0" & CurrentCodeCell.Value
    
    'JAN�����̂܂ܓ����
    ElseIf CurrentCodeCell.Value Like String(13, "#") Then
        
        ForAddinCell.Value = CurrentCodeCell.Value
    
    End If

    '�K�v���ʁA��U�󒍂̐��ʂŖ��߂�B�Z�b�g������ɏ�����������B
    Cells(c.Row, 10).Value = Cells(c.Row, 4).Value

    '���g����
    If c.Value Like "*-*" Then
    
        Call SetParser.ParseScalingSet(c)
    
    End If

Next

End Sub

Private Sub SetParse()
'77777 �Z�b�g����  �s�̑}���𔺂������Ȃ̂ŒP�̂őS���R�[�h�֍s��

Dim ForAddinRange As Range, c As Range
Set ForAddinRange = Range(Cells(2, 9), Cells(Range("B1").SpecialCells(xlCellTypeLastCell).Row, 9))

For Each c In ForAddinRange
    '7777�n�܂�Z�b�g����
    If c.Value Like "7777*" Then

        Call SetParser.ParseItems(c)
    
    End If
    
Next c

'�Z�b�g���i�u�b�N�����
Call SetParser.CloseSetMasterBook

End Sub
