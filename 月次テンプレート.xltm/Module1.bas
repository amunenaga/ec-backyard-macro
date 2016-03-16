Attribute VB_Name = "Module1"
Sub �Ǎ��W�v_�t���I�[�g()
    'MeisaiSheet/paymentMethod��ǂݍ���Ŏ�������āA�V�K�V�[�g�֕ۑ����ďI��
    '�v3�̃t�@�C�����K�v�ł��̂ŁA�}�N���N����̃t�@�C���I���E�B���h�E�Ŏw�肵�ĉ������B
    'CSV:MeisaiSheet,PaymentMethod Xlsx:���i�`�F�b�N����.xlsx
    
    Dim MonthName As String
    MonthName = Format(DateAdd("M", -1, Date), "yy�NM��")
    
    Sheets("���i�ʏW�v").Range("A1") = MonthName & " ���t�[����"
    
    Call PaymentCsv�Ǎ�
    Call MeisaiSheetCsv�Ǎ�
    
    Call �]�L�Əd���폜
    Call �W�v���̑}��
    Call �r��������
    
    Call �������̓]�L
    
    Dim FileName As String
    FileName = "���t�[����" & MonthName & "_��ƒ�.xlsm"
    
    Dim Folder As String
    Folder = Environ("USERPROFILE") & "\Documents\"
    
    Dim Path As String
    Path = Folder & FileName
    
    ThisWorkbook.SaveAs FileName:=Path, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    
    Call ���i�ʏW�v��V�K�t�@�C���փR�s�[
    
End Sub

Private Function findColum(str As String) As Integer

findColum = WorksheetFunction.Match(str, MeisaiSheet.Range("A1").Resize(1, 20), 0)

End Function

Sub MeisaiSheetCsv�Ǎ�()

'Meisai.csv�t�@�C�����w��
'1�s���z��ɓ���āA�V�[�g�֓]�L
'Quantity=0�̓L�����Z���Ȃ̂Œe���܂�

'-----��������-----------------'
Dim FilePath
FilePath = setCsvPath("Meisai")

If FilePath = "" Then
    MsgBox "�t�@�C���w�肪�L�����Z������܂����B"
    Exit Sub
End If

'CSV�Ǎ���TextStream������
Dim LineBuf As Variant
Dim FSO As Object

Set FSO = CreateObject("Scripting.FileSystemObject")

Dim TS As Textstream
Set TS = FSO.OpenTextFile(FilePath, ForReading)
    
'-------�����܂�PaymentMethod�ł��������������Ă��܂�CSV�����Ⴄ����----------"
    
Dim Header As Variant
Header = Split(TS.ReadLine, """,""")

'1���ږڂ�"�ƁA�Ō�̍��ڂ�"���c��̂ō폜���܂��Achr(34)��"�ł�
Header(0) = (Replace(Header(0), Chr(34), ""))
Header(UBound(Header)) = (Replace(Header(UBound(Header)), Chr(34), ""))

'�L�����Z���̒������W�v���Ȃ����߂ɁA���̗�͉��Ԗڂ����肷��
For j = 0 To UBound(Header)
    
    Dim QtyCol As Integer
    
    If Header(j) = "Quantity" Then
        QtyCol = j
        Exit For
    End If

Next

'�w�b�_�[���V�[�g�ɓ]�L
Sheets("Meisai").Range("A1").Resize(1, UBound(Header) + 1).Value = Header

Dim i As Long
i = 1

'������MeisaiSheet�f�[�^���V�[�g�֓]�L
Do Until TS.AtEndOfStream
    
    'LineBuf�z���1���ڂ������
    LineBuf = Split(TS.ReadLine, """,""")
        
    For j = 0 To UBound(LineBuf)
        LineBuf(j) = Trim(Replace(LineBuf(j), Chr(34), "")) '�O�̂��ߍēxchr(34)�� " [���p��d���p��]���������ăg����
        
        If j = QtyCol Then  'qty=0�Ȃ�L�����Z���̒����Ȃ̂ŁA��������Continue�֔��
            If LineBuf(j) = 0 Then GoTo Continue
        
        End If
    
    Next
    
    'A1�Z������A�I�t�Z�b�g�{���T�C�Y���]�L
    Sheets("Meisai").Range("A1").Offset(i, 0).Resize(1, UBound(LineBuf) + 1).Value = LineBuf
    
    i = i + 1

Continue:

Loop

' �w��t�@�C����CLOSE
TS.Close

End Sub

Sub PaymentCsv�Ǎ�()

'Csv���w�肷��
Dim FilePath
FilePath = setCsvPath("PaymentMethod")

If FilePath = "" Then
    MsgBox "�t�@�C���w�肪�L�����Z������܂����B"
    Exit Sub
End If

'CSV�Ǎ��p��TextStream�I�u�W�F�N�g��p��
Dim LineBuf As Variant
Dim FSO As Object

Set FSO = CreateObject("Scripting.FileSystemObject")

Dim TS As Textstream
Set TS = FSO.OpenTextFile(FilePath, ForReading)
    
Dim Header As Variant
Header = Split(TS.ReadLine, """,""")

'1���ږڂ�"�ƁA�Ō�̍��ڂ�"���c��̂ō폜���܂��Achr(34)��"�ł�
Header(0) = (Replace(Header(0), Chr(34), ""))
Header(UBound(Header)) = (Replace(Header(UBound(Header)), Chr(34), ""))

'�w�b�_�[���V�[�g�ɓ]�L
Sheets("PaymentMethod").Range("A1").Resize(1, UBound(Header) + 1).Value = Header

Dim i As Long
i = 1

'������PaymentMethod�̃��R�[�h���V�[�g�֓]�L
Do Until TS.AtEndOfStream
    
    'LineBuf�z���1���ڂ������
    LineBuf = Split(TS.ReadLine, """,""")
        
    For j = 0 To UBound(LineBuf)
        LineBuf(j) = Trim(Replace(LineBuf(j), Chr(34), "")) '�O�̂��ߍēxchr(34)�� " [���p��d���p��]���������ăg����
        
        If j = SaleTotalCol Then  'SaleTotalCol=0�Ȃ�L�����Z���̒����Ȃ̂ŁA��������Continue�֔��
            If LineBuf(j) = 0 Then GoTo Continue
        
        End If
    
    Next
    
    'A1�Z������A�I�t�Z�b�g�{���T�C�Y���]�L
    Sheets("PaymentMethod").Range("A1").Offset(i, 0).Resize(1, UBound(LineBuf) + 1).Value = LineBuf
    
    i = i + 1

Continue:

Loop

' �w��t�@�C����CLOSE
TS.Close

' �ǂݍ��݌�̏W�v ��������ʃv���V�[�W���̕�����������
With PaymentSheet
    
    .Activate '�A�N�e�B�u�łȂ��ƃ_�������A�L���X�g��
    
    Dim EndRow As Long
    EndRow = .Range("A1").End(xlDown).Row
    
    For i = 2 To EndRow
        .Cells(i, 8).NumberFormat = "0"
        .Cells(i, 8).Value = CDbl(Cells(i, 8).Value) 'SUMIF�Ōv�Z����̂ŃZ���̒l�̓_�u���^

    Next

    'Total�̃f�[�^�����W�A�������
    Dim TotalRangeStr As String
    TotalRangeStr = .Rows(1).Find(what:="Total", LookAt:=xlWhole).Offset(1, 0).Resize(EndRow, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    'PaymentMethod�̃f�[�^�����W�A�������
    Dim PaymentMethodRangeStr As String
    PaymentMethodRangeStr = .Rows(1).Find(what:="Payment Method", LookAt:=xlWhole).Offset(1, 0).Resize(EndRow, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    '�W�v�p�̎����Z���֊i�[
        
    .Range("K3").Formula = "=COUNTIF(" & PaymentMethodRangeStr & ",J16)"
    .Range("L3").Formula = "=SUMIF(" & PaymentMethodRangeStr & ",J16," & TotalRangeStr & ")"
    
    .Range("K4").Formula = "=COUNTIF(" & PaymentMethodRangeStr & ",J28)"
    .Range("L4").Formula = "=SUMIF(" & PaymentMethodRangeStr & ",J28," & TotalRangeStr & ")"
    
    .Range("K5").Formula = "=COUNTIF(" & PaymentMethodRangeStr & ",""payment_b1"")"
    .Range("L5").Formula = "=SUMIF(" & PaymentMethodRangeStr & ",""payment_b1""," & TotalRangeStr & ")"
        
    Set AllSaleTotalRange = .Rows(1).Find(what:="", LookAt:=xlPart) '�����E�B���h�E�̐ݒ��߂����߂ɋ󌟍�
        
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


Private Sub �]�L�Əd���폜()
'���㏤�i�̈�ӂȕ\��p�ӂ��܂��B
'MeisaiSheet�̃R�[�h�Ə��i�����W�v�V�[�g�ɓ]�L���ďd���폜���܂��B
'Range���\�b�h��RemoveDuplicates���g���B

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
    
    For i = 2 To MeisaiSheet.UsedRange.Rows.Count
        .Cells(i + 1, 1).Value = MeisaiSheet.Cells(i, DescriptionCol)
        .Cells(i + 1, 2).Value = MeisaiSheet.Cells(i, codeCol)
    Next

    
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

Sub �r��������()
'���i�ʏW�v�̃V�[�g�Ɍr���������܂�

ResizeRow = ItemTotalSheet.Range("A2").End(xlDown).Row - 1

With ItemTotalSheet.Range("A2").Resize(ResizeRow, 9)

    With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With .Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With .Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
End With

End Sub

Sub �������̓]�L()

'�󒍃`�F�b�Nxlsm�t�@�C�����w��
Dim FilePath
FilePath = setCsvPath("���t�[���i�`�F�b�N���w��")

If FilePath = "" Then
    MsgBox "�t�@�C���w�肪�L�����Z������܂����B"
    Exit Sub
End If


'���i�`�F�b�N���[�N�V�[�g�̃R�s�[

Workbooks.Open FilePath
  
With ActiveWorkbook
    
    .Worksheets("���i�`�F�b�N").Copy After:=ThisWorkbook.Worksheets("���i�ʏW�v")
    .Close

End With


'�����`�F�b�N�V�[�g�̏��i�R�[�h���C����Vlookup�Ńq�b�g�����邽�߂�Str�^�Ŋi�[�������B

Worksheets("���i�`�F�b�N").Activate 'With�Ŋ���ƃ����W�w�肪�ʓ|�ɂȂ�̂ŁAActivate���č��
    
Dim EndRow As Integer
EndRow = Worksheets("���i�`�F�b�N").UsedRange.Rows.Count

For i = 2 To EndRow
    Cells(i, 1).NumberFormatLocal = "@"
    Cells(i, 1).Value = CStr(Cells(i, 1).Value)
Next

'Vlookup�Ō�������͈͂��w��
Dim SearchRange As Range
Set SearchRange = Range("A1").Resize(EndRow, 5)


Dim SearchRangeAddress As String
SearchRangeAddress = "���i�`�F�b�N!" & SearchRange.Address(RowAbsolute:=False, ColumnAbsolute:=False)

'Vlookup���𑗂荞��
ItemTotalSheet.Activate

j = 3 '�s�J�E���^�������A�W�v�V�[�g��3�s�ڂ��珤�i�R�[�h���n�܂�

Do Until IsEmpty(Cells(j, 2))
    
    Dim CodeAddress As String
    CodeAddress = Cells(j, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    Cells(j, 2).Offset(0, 5).Formula = "=VLOOKUP(" & CodeAddress & "," & SearchRangeAddress & ",5,FALSE)"
    Cells(j, 2).Offset(0, 6).Formula = "=F" & j & "*" & "G" & j
    
    With Cells(j, 2).Offset(0, 7)
        .Formula = "=H" & j & "/" & "C" & j
        .NumberFormatLocal = "0.00%"
    End With
    
    j = j + 1
    
Loop

Range("H1").Formula = "=SUM(H3:H" & j - 1 & ")"

End Sub

Sub ���i�ʏW�v��V�K�t�@�C���փR�s�[()

ItemTotalSheet.Activate

'����l�ɒ���
'CD FG���3�s�ڂ���ŏI�s�܂ł�l�݂̂ɂ��܂�

'�͈͂�I������
Dim RngCd As Range, RngFg As Range, Rng As Range

Set RngCd = Range("C3:D3").Resize(Range("A1").End(xlDown).Row - 2, 2)
Set RngFg = Range("F3:G3").Resize(Range("A1").End(xlDown).Row - 2, 2)

Set Rng = Union(RngCd, RngFg)

'�����폜���Ēl�݂̂Ƃ���ꍇ�́AValue���i�[�����������ł����I
'http://www.relief.jp/itnote/archives/003686.php�S

For Each c In Rng
    
    c.Value = c.Value

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
