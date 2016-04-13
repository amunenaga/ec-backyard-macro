Attribute VB_Name = "CopyData"
Sub CopyDataBySyokonVendor()
Attribute CopyDataBySyokonVendor.VB_ProcData.VB_Invoke_Func = "r\n14"
'�����}�X�^�[���d�����ʂŃt�B���^�[�\�����āA
'���̕\������Ă���R�[�h�ɂ��āA���t�[�f�[�^�̍s��ʂ̃u�b�N�փR�s�[
'1�d�����ɂ��A1�V�[�g�ɃR�s�[

'�R�[�h���X�g�̏���
SyokonMaster.Activate

'�t�B���^�[�����\���̈�
Dim A As Range
Set A = Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible)

'�R�[�h��̃����W
Dim B As Range
Set B = Range("A2").Resize(Range("A1").SpecialCells(xlCellTypeLastCell).row - 1, 1)

'AB�̌��������W��Code�����W�Ƃ��ăZ�b�g=����d�����̃R�[�h�����W���擾�ł���B
Dim CodeRange As Range
Set CodeRange = Application.Intersect(A, B)


'�R�s�[��u�b�N�̎w��
Dim DestinationBook As Workbook
Set DestinationBook = Workbooks.Add

Dim VendorName As String
VendorName = CodeRange.Cells(1, 4).Value

'�R�s�[��u�b�N�ɐV�����V�[�g��p��
Set NewSheet = DestinationBook.Worksheets.Add()
NewSheet.Name = VendorName

Dim DestinationSheet As Worksheet
Set DestinationSheet = DestinationBook.Worksheets(VendorName)

yahoo6digit.Rows(1).Copy Destination:=DestinationSheet.Rows(1)

SyokonMaster.Activate

Dim r As Range

For Each r In CodeRange
    
    Code = Right(r.Value, 5)
    
    On Error Resume Next
        FindRow = WorksheetFunction.Match(CDbl(Code), yahoo6digit.Range("YahooCodeRange"), 0)
    
    If Err Then
        GoTo continue
    Else
        yahoo6digit.Rows(FindRow).Copy Destination:=DestinationSheet.Rows(DestinationSheet.UsedRange.Rows.Count + 1)
    End If
    
    On Error GoTo 0

continue:

Next

MsgBox VendorName & " �R�s�[����"

End Sub

Sub ExtractYahooData()

'�R�s�[��u�b�N�̎w��
Dim DestinationBook As Workbook
Set DestinationBook = Workbooks.Add

ThisWorkbook.Worksheets("���t�[�f�[�^").Rows(1).Copy Destination:=DestinationBook.Sheets(1).Rows(1)

'���o�������R�[�h���X�g�̗p��
Dim CodeRange As Range
Set CodeRange = Workbooks(2).Sheets(1).Range("B2:B1410")

Dim r As Range

For Each r In CodeRange
    
    Code = r.Value
    
    On Error Resume Next
        FindRow = WorksheetFunction.Match(CDbl(Code), yahoo6digit.Range("YahooCodeRange"), 0)
    
    If Err Then
        r.Interior.ColorIndex = 6
        GoTo continue
    Else
        yahoo6digit.Rows(FindRow).Copy Destination:=DestinationBook.Sheets(1).Rows(DestinationBook.Sheets(1).UsedRange.Rows.Count + 1)
    End If
    
    On Error GoTo 0

continue:

Next

MsgBox VendorName & " �R�s�[����"

End Sub

