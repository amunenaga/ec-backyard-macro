Attribute VB_Name = "Prepare"
Sub ImportYahooCSV()
'���t�[��DataCSV�����t�[�f�[�^�V�[�g�ɃR�s�[���܂��B

'�I�[�g�t�B���^�[������

yahoo6digit.Activate

If Not yahoo6digit.AutoFilter Is Nothing Then yahoo6digit.Range("A1").AutoFilter

Dim DataCsvPath As Variant
' ��t�@�C�����J����̃t�H�[���Ńt�@�C�����̎w����󂯂�
DataCsvPath = Application.GetOpenFilename(Title:="���t�[�̏��i���CSV���w��")

' �L�����Z�����ꂽ�ꍇ�̓��t�[�V�[�g�̍X�V�����ō݌ɎZ�o���s��
If DataCsvPath = False Then

    MsgBox "Yahoo!�V���b�s���O ���i���͍X�V�����ɁA�݌ɂ𐶐����܂��B"
    Range("A1").AutoFilter
    
    Exit Sub

End If

Workbooks.Open DataCsvPath

Dim CsvName As String
CsvName = Dir(DataCsvPath)

'�u���t�[�f�[�^�v���N���A
yahoo6digit.Cells.Clear

Dim RequireHeader As Variant
RequireHeader = Array("path", "name", "code", "price", "sale-price")

With Workbooks(CsvName).Sheets(1)
    '���t�[CSV��XLSM�փR�s�[
    '�w�b�_�[�𒲂ׂĎc���s�ȊO�͍폜
    i = 1
    
    Do Until IsEmpty(.Cells(1, i))
        
        Dim IsReqHeader As Boolean
        IsReqHeader = False
        
        '�K�v�w�b�_�[�Ƃ��ă��X�g�A�b�v���Ă��镶���z��̂ǂ�ł��Ȃ��ꍇ�ɁA��폜
        For Each v In RequireHeader
            If Cells(1, i).Value = v Then
                IsReqHeader = True
            End If
        Next
            
        If IsReqHeader = False Then
            .Columns(i).Delete
        End If
            
            
        i = i + 1
    
    Loop
    
    .Range("A1").CurrentRegion.WrapText = False
    .Range("A1").CurrentRegion.Copy Destination:=yahoo6digit.Range("A1")

    ActiveWindow.Close SaveChanges:=False

End With

End Sub

Sub ImportSyokonAddinData()

Dim ThisFolderPath As String
ThisFolderPath = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "")

Workbooks.Open Filename:=ThisFolderPath & "�����A�h�C���o�̓f�[�^.xlsm"
Application.Run "�����A�h�C���o�̓f�[�^.xlsm!Auto_Open"

SyokonMaster.Cells.Clear

ActiveSheet.Range("A1").CurrentRegion.Copy Destination:=SyokonMaster.Range("A1")

Workbooks("�����A�h�C���o�̓f�[�^.xlsm").Close SaveChanges:=False

End Sub

Sub SetRangeName()
'�e�V�[�g�̃R�[�h�����W���u���O�v�ŌĂׂ�悤�A��`������
'�A�z�z��Ƃ������ăC�e���[�g�񂷗l�ɂ��ׂ�����
'���镔���͊e�X�c�V�[�g���A�ŏ��̃����W�A�����W�� �O���ƃR�s�y���������̕����y��

'���t�[�V�[�g�uYahooCodeRange�v�͈̔͂��Ē�`
With yahoo6digit
    Set rng = .Range("C1").Resize(.Range("C1").SpecialCells(xlCellTypeLastCell).Row, 1)
    Names.Add Name:="YahooCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'�����E�ݔp�́uStockOnlyCodeRange�v�͈̔͂��Ē�`
With StockOnly
    Set rng = .Range("C1").Resize(.Range("C1").SpecialCells(xlCellTypeLastCell).Row, 1)
    Names.Add Name:="StockOnlyCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'�����}�X�^�[�V�[�g�uSyokonCodeRange�v�͈̔͂��Ē�`
With SyokonMaster
    Set rng = .Range("A1").Resize(.Range("A1").SpecialCells(xlCellTypeLastCell).Row, 1)
    Names.Add Name:="SyokonCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'�݌ɃZ�b�g���O�V�[�g
With ExceptQty
    Set rng = .Range("C1").Resize(.Range("C1").SpecialCells(xlCellTypeLastCell).Row, 1)
    Names.Add Name:="ExceptCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'�p�ԃV�[�g�uEolCodeRange�v
With Eol
    Set rng = .Range("C1").Resize(.Range("C1").SpecialCells(xlCellTypeLastCell).Row, 1)
    Names.Add Name:="EolCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

'SLIMS�V�[�g�uSlimsCodeRange�v
With Slims
    Set rng = .Range("B1").Resize(.Range("B1").SpecialCells(xlCellTypeLastCell).Row, 1)
    Names.Add Name:="SlimsCodeRange", RefersTo:="=" & .Name & "!" & rng.Address
End With

End Sub

