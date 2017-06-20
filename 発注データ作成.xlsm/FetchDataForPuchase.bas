Attribute VB_Name = "FetchDataForPuchase"
Option Explicit

Sub CreateQuantitySheet()

Call LoadPurchaseReq.LoadAllPicking

Worksheets("��z���ʌ���V�[�g").Activate

Call SumPuchaseRequest

'�����ɕK�v�ȏ����f�[�^�x�[�X�EExcel�t�@�C������擾
Call FetchSyokonData
Call FetchExcellForPurchase

Call CalcPurchaseQuantity

Call FetchPickupFlag

Call FetchExcellJanInventory

End Sub

Sub FetchSyokonData()
'�������i�}�X�^�[�̃f�[�^�擾
'���ۂɓǂ݂ɍs���e�[�u���́A�󒍏ڍ׊m�F�p�ɖ������v���P�[�V���������Server3��EC�ۗp�}�X�^

'�ڑ��̂��߂̃I�u�W�F�N�g���`�ADB�ڑ��ݒ���Z�b�g
Dim DbCnn As New ADODB.Connection
Dim DbCmd  As New ADODB.Command
Dim DbRs As New ADODB.Recordset

DbCnn.ConnectionTimeout = 0
DbCnn.Open "PROVIDER=SQLOLEDB;Server=Server02;Database=ITOSQL_REP;UID=sa;PWD=;"
DbCmd.CommandTimeout = 180
Set DbCmd.ActiveConnection = DbCnn

'���i�R�[�h�̃����W���Z�b�g�A1�Z������SQL���s
Dim CodeRange As Range, r As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange
    Dim sql As String, Code As String
    Code = r.Value
        
    sql = "SELECT ���i�R�[�h, �戵�敪, ���b�g��, �d������, �d����, �d����}�X�^.�d���旪��, �d����}�X�^.�����敪 " & _
          "FROM ���i�}�X�^ JOIN �d����}�X�^ ON ���i�}�X�^.�d���� = �d����}�X�^.�d����R�[�h " & _
          "WHERE ���i�R�[�h = " & Code & "OR JAN�R�[�h = '" & Code & "'"
    
    Set DbRs = DbCnn.Execute(sql)

    If Not DbRs.EOF Then
        Cells(r.Row, 3).Value = DbRs("���b�g��")
        Cells(r.Row, 4).Value = DbRs("�d����")
        Cells(r.Row, 5).Value = DbRs("�d���旪��")
        Cells(r.Row, 10).Value = DbRs("�d������")
        Cells(r.Row, 2).Value = GetKubunLabel(DbRs("�戵�敪"))
        Cells(r.Row, 11).Value = DbRs("�����敪")
        
        'JAN�󒍕��̏��i�R�[�h�u���A���Amazon���p
        If Len(r.Value) > 6 Then
            r.NumberFormatLocal = "@"
            r.Value = IIf(Len(DbRs("���i�R�[�h")) = 5, "0" & DbRs("���i�R�[�h"), DbRs("���i�R�[�h"))
        End If
    
    End If

Next

End Sub

Sub FetchExcellForPurchase()
'�����p���i���̃f�[�^�擾

Dim DataBook As Workbook, DataSheet As Worksheet, PurDataCodeRange As Range, PurDataJanRange As Range
Set DataSheet = OpenPurDataBook().Worksheets("���i���")
Set PurDataJanRange = DataSheet.Range(Cells(1, 1), Cells(DataSheet.UsedRange.Rows.Count, 1))
Set PurDataCodeRange = DataSheet.Range(Cells(1, 2), Cells(DataSheet.UsedRange.Rows.Count, 2))

ThisWorkbook.Activate
Dim CodeRange As Range, r As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange

    Dim Code As String, HitRow As Double
        
    Code = r.Value

    On Error Resume Next
        HitRow = WorksheetFunction.Match(Code, PurDataCodeRange, 0)

        If Err Then
            Err.Clear
            HitRow = WorksheetFunction.Match(Code, PurDataJanRange, 0)
            
            If Err Or IsEmpty(Cells(r.Row, 4).Value) Then '�d����R�[�h�������ɂ��G�N�Z���ɂ��Ȃ���΁A�����ł��Ȃ��̂Œ��ӏ���������
                Cells(r.Row, 2).Value = "�����p���i��� �f�[�^�Ȃ�"
                GoTo Continue
            End If
        
        End If
    
    On Error GoTo 0
        
    '��z�����ӁA���[�J�[���b�g�A�d���於
    Cells(r.Row, 2).Value = Cells(r.Row, 2).Value & DataSheet.Cells(HitRow, 35).Value '��z������
    Cells(r.Row, 12).Value = DataSheet.Cells(HitRow, 5).Value '�����p���i���̃��b�g��
    Cells(r.Row, 13).Value = DataSheet.Cells(HitRow, 4).Value '�d���於
    
    '�d����R�[�h�A�����A�d���於��6�P�^�ɂȂ����̂ݓ����
    If IsEmpty(Cells(r.Row, 4).Value) Then
    
        Cells(r.Row, 4).Value = DataSheet.Cells(HitRow, 32).Value '�d����R�[�h
        Cells(r.Row, 5).Value = DataSheet.Cells(HitRow, 4).Value '�d���於
        Cells(r.Row, 10).Value = DataSheet.Cells(HitRow, 13).Value '����

    End If

Continue:

Next

End Sub

Sub FetchExcellJanInventory()
'�I�Ȃ��݌ɕ\�f�[�^�̊m�F

Const NOLOCATION_INVENTRY_EXCELL As String = "\\server02\���i��\�l�b�g�̔��֘A\�I���݌Ɋm�F�\.xlsm"
Const INVENTRY_SHEET As String = "�I���f�[�^"

Workbooks.Open FileName:=NOLOCATION_INVENTRY_EXCELL, ReadOnly:=True

Dim CodeRange As Range, r As Range
With ThisWorkbook.Worksheets("��z���ʌ���V�[�g")
    Set CodeRange = .Range(.Cells(2, 7), .Cells(2, 7).End(xlDown))
End With

Dim InventryRange As Range

With Workbooks(Dir(NOLOCATION_INVENTRY_EXCELL)).Worksheets(INVENTRY_SHEET)
    
    Set InventryRange = .Range(.Cells(1, 2), .Cells(1, 2).End(xlDown))

End With

For Each r In CodeRange

    Dim Code As String, HitRow As Double, Location As String, StockQuantity As Long
    
    Code = r.Value

    On Error Resume Next
    HitRow = WorksheetFunction.Match(Code, InventryRange, 0)
    
    If Err = 0 Then
    
        StockQuantity = InventryRange.Cells(HitRow, 1).Offset(0, 1).Value
            
        If StockQuantity > 0 Then
            
            Location = InventryRange.Cells(HitRow, 1).Offset(0, 3).Value
            r.Offset(0, -5).Value = "�I��:" & StockQuantity & "�ꏊ:" & Location
        
        End If
            
    End If
    
    On Error GoTo 0

Next

Workbooks(Dir(NOLOCATION_INVENTRY_EXCELL)).Close Savechanges:=False

End Sub

Sub CalcPurchaseQuantity()
'��z�˗����ʂ���A���b�g�P�ʁE�����P�ʂŊۂ߂����������ʂ��Z�o���AA��֓����B

Dim CodeRange As Range, r As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange
    Dim i As Long
    i = r.Row
    
    Dim Rot As Double, Qty As Long, RequestQty As Double
    Rot = Cells(i, 12).Value
    
    If IsEmpty(Rot) Or Rot = 0 Then
        Rot = 1
    End If
    
    RequestQty = Cells(i, 9).Value
    
    Qty = WorksheetFunction.Ceiling(RequestQty, Rot)

    Cells(i, 1).Value = Qty
Next

End Sub

Private Function GetKubunLabel(ByVal KubunCode As Integer) As String
'���i�}�X�^�ł͋敪��1�`9�̐����Ȃ̂ŁA�\�����Œu��������B
'���̃}�X�^���ɐ���-�敪���̑g�͕ۑ�����Ă��邪�A�����ł�Switch���ŐU�蕪����B

Dim tmp As String

Select Case KubunCode
    Case 3
        tmp = "����:�̔����~"
    Case 7
        tmp = "����:�݌ɔp��"
    Case 8
        tmp = "����:�݌ɏ���"
    Case 9
        tmp = "����:���[�J�[�p��"
    Case Else
        tmp = ""
End Select

GetKubunLabel = tmp

End Function

Private Function OpenPurDataBook() As Workbook

'�����p���i���̃t�@�C�����J���܂��B
'���s���̃G�N�Z���Ŕ����p���i���̃t�@�C�����J���Ă���΁A���̃��[�N�u�b�N��Ԃ��܂��B

Const PUR_DATA_EXCELL_PATH As String = "\\Server02\���i��\�l�b�g�̔��֘A\�����֘A\�����p���i���.xlsm"

Dim WorkBookName As String
WorkBookName = Dir(PUR_DATA_EXCELL_PATH)

Dim wb As Workbook

For Each wb In Workbooks
    
    If wb.Name = WorkBookName Then
        wb.Activate
        GoTo ret
    
    End If

Next

Set wb = Workbooks.Open(PUR_DATA_EXCELL_PATH)

ret:

Set OpenPurDataBook = wb

End Function

Sub FetchPickupFlag()

'�ڑ��̂��߂̃I�u�W�F�N�g���`�ADB�ڑ��ݒ���Z�b�g
Dim DbCnn As New ADODB.Connection
Dim DbCmd  As New ADODB.Command
Dim DbRs As New ADODB.Recordset

DbCnn.ConnectionTimeout = 0
DbCnn.Open "PROVIDER=SQLOLEDB;Server=Server02;Database=ITOSQL_REP;UID=sa;PWD=;"
DbCmd.CommandTimeout = 180
Set DbCmd.ActiveConnection = DbCnn

'�d����R�[�h�̃����W���Z�b�g�A1�Z������SQL���s
Dim CodeRange As Range, r As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown)).Offset(0, -3)

For Each r In CodeRange

    '����敪���󗓂ŁA�d����R�[�h�������Ă���Ώ�������擾
    If Cells(r.Row, 11).Value = "" And Not Cells(r.Row, 4).Value = "" Then
    
    Dim sql As String, VendorCode As String
    VendorCode = r.Value
        
    sql = "SELECT �����敪 " & _
          "FROM �d����}�X�^ " & _
          "WHERE �d����R�[�h = " & VendorCode
    
    Set DbRs = DbCnn.Execute(sql)

    If Not DbRs.EOF Then
        Cells(r.Row, 11).Value = DbRs("�����敪")
    End If

    End If

Next


End Sub
