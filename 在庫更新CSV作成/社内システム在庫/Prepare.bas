Attribute VB_Name = "Prepare"
Sub FetchYahooCSV()
'���t�[��DataCSV�����t�[�f�[�^�V�[�g�ɃR�s�[���܂��B

'�I�[�g�t�B���^�[������

yahoo6digit.Activate

If Not yahoo6digit.AutoFilter Is Nothing Then yahoo6digit.Range("A1").AutoFilter

'�u���t�[�f�[�^�v���N���A
yahoo6digit.Cells.Clear

Dim DataCsvPath As String
' ��t�@�C�����J����̃t�H�[���Ńt�@�C�����̎w����󂯂�
DataCsvPath = Application.GetOpenFilename(Title:="���t�[�̏��i���CSV���w��")

' �L�����Z�����ꂽ�ꍇ��False���Ԃ�̂ňȍ~�̏����͍s�Ȃ�Ȃ�
If VarType(DataCsvPath) = vbBoolean Then Exit Sub

Workbooks.Open DataCsvPath

Dim CsvName As String
CsvName = Dir(DataCsvPath)

Dim header As Variant
header = Array("sub-code", "original-price", "options", "caption")  '"headline"

With Workbooks(CsvName).Sheets(1)

    '���t�[CSV��XLSM�փR�s�[
    '�w�b�_�[�𒲂ׂ�Abstract�܂ł̊ԂɁAsub-code/original-price/options/headline/caption������Η���폜
    i = 1
    Do Until IsEmpty(.Cells(1, i))
        For Each v In header
            If Cells(1, i) = v Then
                .Columns(i).Delete
            End If
        Next
            
        i = i + 1
    
    Loop
    
    .Range("A1").CurrentRegion.WrapText = False
    .Range("A1").CurrentRegion.Copy Destination:=yahoo6digit.Range("A1")

    ActiveWindow.Close saveChanges:=False

End With

End Sub


Sub FetchSecondInventry()

'�I���݌Ɋm�F�\���J���ĒI���f�[�^=SecondInventry�ɃR�s�[

Application.ScreenUpdating = False

Const SECOND_INVENTRY_FILE As String = "\\server02\���i��\�l�b�g�̔��֘A\�I���݌Ɋm�F�\.xlsm"
Const SECOND_INVENTRY_SHEET_NAME As String = "�I���f�[�^"

SecondInventry.Cells.Clear

'�݌ɕ\���J���ăV�[�g���R�s�[
'1.�݌ɕ\�̑��݃`�F�b�N

Dim WbName As String
WbName = Dir(SECOND_INVENTRY_FILE)

If WbName = "" Then
    MsgBox "�I�����̍݌ɕ\�����݂��܂���", vbExclamation
    Exit Sub
End If

'2.�����u�b�N���J���Ă��Ȃ����`�F�b�N
Dim wb As Workbook

For Each wb In Workbooks
    If wb.Name = WbName Then
        MsgBox WbName & vbCrLf & "�͂��łɊJ���Ă��܂�", vbExclamation
        Exit Sub
    End If
Next wb

'�����Ńu�b�N���J��
Workbooks.Open SECOND_INVENTRY_FILE

'�݌ɕ\���V�[�g���R�s�[
For i = 1 To Workbooks.Count
    
    If Workbooks(i).Name = WbName Then
        
        With Workbooks(i).Sheets(SECOND_INVENTRY_SHEET_NAME)
            
            Dim LastRow As Long
            LastRow = .Range("A1").SpecialCells(xlCellTypeLastCell).row
            
            .Range("A1").Resize(LastRow - 1, 4).Copy
        
        End With
        
        SecondInventry.Range("A1").PasteSpecial (xlPasteValues)
        
        Application.DisplayAlerts = False
        
        '�R�s�[���I���Α��₩�ɍ݌ɕ\�����
        Workbooks(WbName).Close saveChanges:=False
        
        Application.DisplayAlerts = True
        
        Exit For
        
    End If

Next

Worksheets(SECOND_INVENTRY_SHEET_NAME).Range("A1").AutoFilter

With Worksheets(SECOND_INVENTRY_SHEET_NAME).AutoFilter.Sort

    .SortFields.Clear '�\�[�g�t�B�[���h����U�N���A�[
    
    '�\�[�g�t�B�[���h���w��
    .SortFields.Add Key:=Range("C1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    
    '�\�[�g�����w��
    .header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    
    '�\�[�g�K�p
    .Apply

End With

End Sub

Sub SetRangeName()
'�e�V�[�g�̃R�[�h�����W���u���O�v�ŌĂׂ�悤�A��`������
'�A�z�z��Ƃ������ăC�e���[�g�񂷗l�ɂ��ׂ�����
'���镔���͊e�X�c�V�[�g���A�ŏ��̃����W�A�����W�� �O���ƃR�s�y���������̕����y��

'���t�[�V�[�g�uYahooCodeRange�v�͈̔͂��Ē�`
With yahoo6digit
    Set rng = .Range("C1").Resize(.Range("C1").SpecialCells(xlCellTypeLastCell).row, 1)
    Names.Add Name:="YahooCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'�����E�ݔp�́uStockOnlyCodeRange�v�͈̔͂��Ē�`
With StockOnly
    Set rng = .Range("C1").Resize(.Range("C1").SpecialCells(xlCellTypeLastCell).row, 1)
    Names.Add Name:="StockOnlyCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'�����}�X�^�[�V�[�g�uSyokonCodeRange�v�͈̔͂��Ē�`
With SyokonMaster
    Set rng = .Range("A1").Resize(.Range("A1").SpecialCells(xlCellTypeLastCell).row, 1)
    Names.Add Name:="SyokonCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'�݌ɃZ�b�g���O�V�[�g
With ExceptQty
    Set rng = .Range("C1").Resize(.Range("C1").SpecialCells(xlCellTypeLastCell).row, 1)
    Names.Add Name:="ExceptCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'�p�ԃV�[�g�uEolCodeRange�v
With Eol
    Set rng = .Range("C1").Resize(.Range("C1").SpecialCells(xlCellTypeLastCell).row, 1)
    Names.Add Name:="EolCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'�I���݌ɕ\�V�[�g�uSecondInventryCodeRange�v
With SecondInventry
    Set rng = .Range("B1").Resize(.Range("B1").SpecialCells(xlCellTypeLastCell).row, 1)
    Names.Add Name:="SecondInventryCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

End Sub

