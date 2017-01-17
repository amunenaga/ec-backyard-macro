Attribute VB_Name = "Module1"
Sub �Ǎ��W�v_�t���I�[�g()
    'MeisaiSheet/paymentMethod��ǂݍ���Ŏ�������āA�V�K�V�[�g�֕ۑ����ďI��
    'CSV:MeisaiSheet,PaymentMethod
    
    Dim MonthName As String
    MonthName = Format(DateAdd("M", -1, Date), "yy�NM��")
    
    Sheets("���i�ʏW�v").Range("A1") = MonthName & " ���t�[����"
    
    Call meisaiCSV�C���|�[�g
    
    Call �]�L�Əd���폜
    Call �W�v���̑}��
    Call �r��������
    
    Dim FileName As String
    FileName = "���t�[����" & MonthName & "_��ƒ�.xlsm"
    
    Dim Folder As String
    Folder = Environ("USERPROFILE") & "\Documents\"
    
    Dim Path As String
    Path = Folder & FileName
    
    ThisWorkbook.SaveAs FileName:=Path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    
    Call ���i�ʏW�v��V�K�V�[�g�փR�s�[
    
End Sub

Private Function findColum(str As String) As Integer

findColum = WorksheetFunction.Match(str, MeisaiSheet.Range("A1").Resize(1, 20), 0)

End Function

Sub meisaiCSV�C���|�[�g()

Dim FilePath
FilePath = setCsvPath("Meisai")

If FilePath = "" Then
    MsgBox "�t�@�C���w�肪�L�����Z������܂����B"
    Exit Sub
End If

MeisaiSheet.Activate

With ActiveSheet.QueryTables.Add(Connection:= _
    "TEXT;" & FilePath, Destination:=Range("$A$1"))
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
    .TextFileCommaDelimiter = True
    .TextFileSpaceDelimiter = False
    .TextFileColumnDataTypes = Array(2, 1, 1, 2, 2, 1, 1, 1, 1, 1, 1, 1)
    .TextFileTrailingMinusNumbers = True
    .Refresh BackgroundQuery:=False
End With

End Sub

Private Function setCsvPath(CsvName As String)
'�t�@�C���I���_�C�A���O���J���ăt�@�C���w��A�p�X��Ԃ�

' ��t�@�C�����J����̃t�H�[���Ńt�@�C�����̎w����󂯂�
Path = Application.GetOpenFilename(Title:=CsvName & "���w��")

' �L�����Z�����ꂽ�ꍇ��False���Ԃ�̂ňȍ~�̏����͍s�Ȃ�Ȃ�
If VarType(Path) = vbBoolean Then Exit Function

setCsvPath = Path
    
End Function

Sub �]�L�Əd���폜()
'���㏤�i�̈�ӂȕ\��p�ӂ��܂��B
'MeisaiSheet�̃R�[�h�Ə��i�����W�v�V�[�g�ɓ]�L���ďd���폜���܂��B
'Range���\�b�h��RemoveDuplicates���g���B

ItemTotalSheet.Activate

'���i�ʏW�v�V�[�g�ւ̓]�L

On Error GoTo ErrorMes
   'Code��̓���
    Dim codeCol As Integer
    codeCol = WorksheetFunction.Match("Product Code", MeisaiSheet.Range("A1").Resize(1, 20), 0)
    
    'Description��̓���
    Dim DescriptionCol As Integer
    DescriptionCol = WorksheetFunction.Match("Description", MeisaiSheet.Range("A1").Resize(1, 20), 0)

On Error GoTo 0


'���i�ʏW�v��Meisai�V�[�g����]�L
'Code=���i�R�[�h��Description=���i����1�s����

With ItemTotalSheet
    
    Dim i As Long
    i = 2

    Do While MeisaiSheet.Cells(i, codeCol).Value <> ""
        
        '����0�̓L�����Z���Ȃ̂Ŕ�΂�
        If MeisaiSheet.Cells(i, 3).Value = 0 Then GoTo Continue
        
        Dim WriteRow As Long
        WriteRow = .Range("A1").SpecialCells(xlCellTypeLastCell).Row + 1
        
        .Cells(WriteRow, 1).Value = MeisaiSheet.Cells(i, DescriptionCol)
        .Cells(WriteRow, 2).Value = MeisaiSheet.Cells(i, codeCol)
        
Continue:
    i = i + 1
    Loop
    
    'Range���w�肵��Range�I�u�W�F�N�g��RemoveDuplicate���\�b�h�ňꔭ�d���폜��G�N�Z��2010�ȍ~�B
    
    .Range("A2:B2").Resize(.UsedRange.Rows.Count, 2).Name = "���i���X�g"
    .Range("���i���X�g").RemoveDuplicates Columns:=2, Header:=xlYes

End With


Exit Sub

ErrorMes:
MsgBox "�����𒆎~���܂����B" & vbLf & "Meisai�V�[�g��Product Code��Description������܂���B"

End Sub

Sub �W�v���̑}��()

'���i�R�[�h�̍ŏ��̃Z������ŏI�s�܂ł�Range���i�[
'SUM�֐����g�����߂ɁA���l�^�ɕϊ����K�v�ȃ����W�����킹�Ċi�[���Ă����B

Dim sh1EndRow As Long
sh1EndRow = MeisaiSheet.UsedRange.Rows.Count

'���v�A���A�����ȂǏW�v�Ώۂ̗���_�u���^�ɃL���X�g
  
Dim Rng As Range
Set Rng = Union(MeisaiSheet.Cells(2, findColum("Quantity")).Resize(sh1EndRow, 1), _
                MeisaiSheet.Cells(2, findColum("Unit Price")).Resize(sh1EndRow, 1), _
                MeisaiSheet.Cells(2, findColum("Line Sub Total")).Resize(sh1EndRow, 1))

Dim c As Range
For Each c In Rng
    c.NumberFormat = "General" '�\���`�����u�W���v�ɃZ�b�g
    c.Value = CDbl(c.Value)    '�_�u���^�ɃL���X�g���Ċi�[
Next

'Code��̓���
On Error GoTo ErrorMes
    Dim codeCol As Integer
    codeCol = WorksheetFunction.Match("Product Code", MeisaiSheet.Range("A1").Resize(1, 20), 0)
    
On Error GoTo 0

Dim MeisaiCodeRange As Range
Set MeisaiCodeRange = MeisaiSheet.Cells(2, codeCol).Resize(sh1EndRow - 1, 1)

i = 3

With ItemTotalSheet
    Do Until IsEmpty(.Cells(i, 2))
        
        '���i�R�[�h�ɑ΂��锄����z�����v����SUMIF���A
        .Cells(i, "C").Formula = "=SUMIF(Meisai!" & MeisaiCodeRange.Address & ",B" & i & ",Meisai!" & MeisaiCodeRange.Offset(0, 6).Address & ")"
            
        '���i�R�[�h�ɑ΂��钍�����������v����COUNTIF���A
        .Cells(i, "D").Formula = "=COUNTIF(Meisai!" & MeisaiCodeRange.Address & ",B" & i & ")"
    
        '���i�R�[�h�ɑ΂��锄��������v����SUMIF��
        .Cells(i, "F").Formula = "=SUMIF(Meisai!" & MeisaiCodeRange.Address & ",B" & i & ",Meisai!" & MeisaiCodeRange.Offset(0, -1).Address & ")"
        
        .Cells(i, "E").Formula = "=C" & i & "/F" & i '���Z������̂ōŌ�ɍs��
        
    i = i + 1
    
    Loop
    
    Dim EndRow As Integer
    EndRow = .Range("C3").End(xlDown).Row
    
    .Range("C1").Formula = "=SUM(C3:C" & EndRow & ")"
    .Range("H1").Formula = "=SUM(H3:H" & EndRow & ")"

End With

Exit Sub

ErrorMes:

MsgBox "�����𒆎~���܂����B" & vbLf & "Meisai�V�[�g��Product Code��Description������܂���B"

End Sub

Private Sub �r��������()
'���i�ʏW�v�̃V�[�g�Ɍr���������܂�

ResizeRow = ItemTotalSheet.Range("A2").End(xlDown).Row - 1

With ItemTotalSheet.Range("A2").Resize(ResizeRow, 9).Borders
        
        .LineStyle = xlContinuous
        .Weight = xlThin
    
End With

End Sub

Private Sub ���i�ʏW�v��V�K�t�@�C���փR�s�[()

ItemTotalSheet.Activate

'����l�ɒ���
'CD FG���3�s�ڂ���ŏI�s�܂ł�l�݂̂ɂ��܂�

'�͈͂�I������
Dim RngCd As Range, RngFg As Range, Rng As Range

Set RngCd = Range("C3:D3").Resize(Range("A1").End(xlDown).Row - 2, 2)
Set RngFg = Range("F3:G3").Resize(Range("A1").End(xlDown).Row - 2, 2)

Set Rng = Union(RngCd, RngFg)

'�����폜���Ēl�݂̂Ƃ���ꍇ�́AValue���i�[�����������ł���
'http://www.relief.jp/itnote/archives/003686.php

For Each c In Rng
    
    c.Value = c.Value

Next

'�R�[�h��6�P�^�ɏC��
Dim RngCode As Range
Set RngCode = Range("B3").Resize(Range("A1").End(xlDown).Row - 2, 1)

For Each c In RngCode
    
    c.NumberFormatLocal = "@"
    If Len(c.Value) = 5 Then c.Value = "0" & c.Value

Next

'�V�K���[�N�u�b�N�֏��i�ʏW�v���R�s�[
Dim FileName As String, FolderPath As String, Path As String

FolderPath = Environ("USERPROFILE") & "\Documents\"
FileName = "���t�[����" & Format(DateAdd("M", -1, Date), "yy�NM��") & ".xlsx"

Path = FolderPath & FileName

Sheets("���i�ʏW�v").Copy

ActiveWorkbook.SaveAs Path

ThisWorkbook.Close SaveChanges:=False

End Sub

Private Sub ���i�ʏW�v��V�K�V�[�g�փR�s�[()

ItemTotalSheet.Activate
ItemTotalSheet.Copy After:=Worksheets("���i�ʏW�v")

ActiveSheet.Name = "��������"

'����l�ɒ���
'CD FG���3�s�ڂ���ŏI�s�܂ł�l�݂̂ɂ��܂�

'�͈͂�I������
Dim RngCd As Range, RngFg As Range, Rng As Range

Set RngCd = Range("C3:D3").Resize(Range("A1").End(xlDown).Row - 2, 2)
Set RngFg = Range("F3:G3").Resize(Range("A1").End(xlDown).Row - 2, 2)

Set Rng = Union(RngCd, RngFg)

'�����폜���Ēl�݂̂Ƃ���ꍇ�́AValue���i�[�����������ł���
'http://www.relief.jp/itnote/archives/003686.php

For Each c In Rng
    
    c.Value = c.Value

Next

'�R�[�h��6�P�^�ɏC��
Dim RngCode As Range
Set RngCode = Range("B3").Resize(Range("A1").End(xlDown).Row - 2, 1)

For Each c In RngCode
    
    c.NumberFormatLocal = "@"
    If Len(c.Value) = 5 Then c.Value = "0" & c.Value

Next

ActiveSheet.Shapes(1).Delete

End Sub
