Attribute VB_Name = "BuildShippingNumSheet"
Option Explicit
Sub ���샄�}�g_�V�[�g�쐬()

'�t�@�C���ۑ�
Application.DisplayAlerts = False
    ThisWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\�o�׊m�F_" & Format(Date, "yyyyMMdd") & ".xlsx", FileFormat:=xlWorkbookDefault
Application.DisplayAlerts = True


'TSV/CSV�t�@�C���Ǎ�
Call LoadAllCsv

'�{�^���폜
Worksheets("�g�b�v").Shapes(1).Delete

'�f�[�^�擾�㏈��  �f�[�^�����N�폜���Z���́u���O�v�폜
Dim qt As QueryTable
For Each qt In Worksheets("�g�b�v").QueryTables
    qt.Delete
Next qt

Dim nm As Name
For Each nm In ActiveWorkbook.Names
    nm.Delete
Next nm

'�^����ЕʂɃV�[�g�փR�s�[
'�^����Ж��̑����ԍ��`��5�P�^�́ASortByCarrier�v���V�[�W���ɂăn�[�h�R�[�f�B���O
Call SortByCarrier("����}��")
Call SortByCarrier("���}�g�^�A")

'�񕝒���
Dim i As Long
For i = 1 To Worksheets.Count
    Worksheets(i).Range("A1").CurrentRegion.Columns.AutoFit
Next i

'�U����̕ۑ��Ɗ������b�Z�[�W
ThisWorkbook.Save

MsgBox "����"
End Sub

Private Sub LoadAllCsv()
'�����ԍ�CSV�p�X���擾�A3�t�@�C���܂œ����w��\

'�t�@�C���_�C�A���O�ɂ�Amazon�E�y�V�E���t�[��TSV/CSV�t�@�C���𕡐������w�肵�Ă��炤
Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd
    '�t�@�C���I���_�C�A���O�̐ݒ�
    .Filters.Clear
    .Filters.Add "Amazon,�y�V,Yahoo!", "*.tsv; *.csv"
    .InitialFileName = "\\Server02\���i��\�l�b�g�̔��֘A\����f�[�^\�o�גʒm"
    
    '�_�C�A���O�\��
    .Show
    
    '�t�@�C���I����̏���
    If .SelectedItems.Count = 0 Then
    
        MsgBox "�t�@�C���w�肪�L�����Z������܂����B"
        End
    
    ElseIf .SelectedItems.Count >= 4 Then
        MsgBox "�t�@�C���w�肪3�𒴂��Ă��܂��B"
        End
    
    End If
    
    Dim Paths(2) As String, CurrentPath As String, i As Long
    
    '�I�����ꂽ�t�@�C���p�X���璆�g���ׂă��[�����ɃZ�b�g
    For i = 1 To .SelectedItems.Count
        Call LoadCsv(.SelectedItems.Item(i))
    Next
    
End With

End Sub

Private Sub LoadCsv(ByVal Path As String)
'�����̃p�X�Ƀe�L�X�g�X�g���[���Őڑ����ăw�b�_�[�𒲂ׁA���[������Ԃ��B
Dim FSO As Object, TS As Object, i As Long, CurrentMall As String, CurrentRow As Variant

Set FSO = CreateObject("Scripting.FileSystemObject")
Set TS = FSO.OpenTextFile(Path)
        
Do Until TS.AtEndOfStream Or i > 3
    CurrentRow = TS.ReadLine
    
    '�^�u�������Amazon
    If InStr(CurrentRow, Chr(9)) > 0 Then
        Call LoadAmazon(Path)
        Exit Do
    
    '�󒍔ԍ� �̕���������Ίy�V
    ElseIf InStr(CurrentRow, "�󒍔ԍ�") > 0 Then
        Call LoadRakuten(Path)
        Exit Do
        
    'OrderId �̕���������΃��t�[
    ElseIf InStr(CurrentRow, "OrderId") > 0 Then
        Call LoadYahoo(Path)
        Exit Do
    
    End If
    
    i = i + 1

Loop

End Sub

Sub SortByCarrier(ByVal CarrierName As String)
'�^����Ж����󂯎���āA�^����Ж��̃V�[�g�֑����ԍ����R�s�[
'�����ŃR�s�[��V�[�g���w�肷��̂ŁA�V�[�g����Case�\���̉^����Ђƍ��킹�邱�ƁB

Dim Criteria As Variant, Operator As Integer

Worksheets("�g�b�v").Activate

'����̑����ԍ��`��4�P�^�A�t�B���^�[�̃v���p�e�B�ɂȂ�̂Ŕz��֑��
Criteria = Array("4031*", "4012*")

Select Case CarrierName
    
    Case "����}��"
        'Criteria�z��ɓ����Ă��鑗���ԍ��Ńt�B���^�[
        Operator = xlFilterValues
        
        With Range("A1").CurrentRegion
            .AutoFilter Field:=3, Criteria1:=Criteria, Operator:=Operator
            .Copy Worksheets(CarrierName).Range("A1")
            .AutoFilter '�I�[�g�t�B���^�[����
        End With
        
    Case "���}�g�^�A"
        '����ȊO�̑����ԍ����t�B���^�[�ACriteria��2�܂ł����Z�b�g�o���Ȃ��B
        
        Operator = xlAnd
        
        With Range("A1").CurrentRegion
            .AutoFilter Field:=3, Criteria1:="<>" & Criteria(0), Operator:=Operator, Criteria2:="<>" & Criteria(1)
            .Copy Worksheets(CarrierName).Range("A1")
            .AutoFilter '�I�[�g�t�B���^�[����
        End With
End Select

End Sub

Sub LoadAmazon(ByVal Path As String)

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & Path, Destination:=GetDestRange())
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
    "TEXT;" & Path, Destination:=GetDestRange())
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
    "TEXT;" & Path, Destination:=GetDestRange()) '�p�X�Ə����o����͓��I�Ɍ��߂�
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

Private Function GetDestRange() As Range

'�����o����Z�������߂�AB2����̎���End�R�}���h�ł�1,048,576�s�܂Ŕ��ł��܂��̂ŁB
Dim r As Range
If IsEmpty(Range("B2")) Then
    Set r = Range("B2")
Else
    Set r = Range("B1").End(xlDown).Offset(1, 0)
End If

Set GetDestRange = r

End Function
