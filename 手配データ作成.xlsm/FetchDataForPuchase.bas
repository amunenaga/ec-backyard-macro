Attribute VB_Name = "FetchDataForPuchase"
Option Explicit

Sub CreateQuantitySheet()
'�s�b�L���O�V�[�g�����z�˗�����ǂݍ���ŁA���i�ʂɏW�v�A�d����f�[�^�Ȃǂ���������Ǎ�
'�u��z�����̓V�[�g�쐬�v�{�^���ŌĂяo�����

Application.ScreenUpdating = False
Application.DisplayAlerts = False

'�Z���[���A�����A��z���ʓ��̓V�[�g��p��
Dim Sh As Variant
For Each Sh In Array(Worksheets("�Z���[��"), Worksheets("����"), Worksheets("��z���ʓ��̓V�[�g"))
    Call PrepareSheet(Sh)
Next

'�A�}�]���E�y�V�E���t�[�̊e�I�Ȃ��s�b�L���O�V�[�g�A�A�}�]�����̎�z�˗��Ǎ�
Call LoadPurchaseReq.LoadAllPicking

ThisWorkbook.SaveAs FileName:=ThisWorkbook.path & "\" & "��z�f�[�^" & Format(Date, "MMdd") & ".xlsm"

Worksheets("��z���ʓ��̓V�[�g").Activate

'���i�ʂɎ�z�˗����ʂ��W�v
Call SumPuchaseRequest

'�����ɕK�v�ȏ����f�[�^�x�[�X�EExcel�t�@�C������擾
Call FetchSyokonData
Call FetchExcellForPurchase

Call CalcPurchaseQuantity

Call FetchPickupFlag

'Excel�ŊǗ�����Ă���I�Ȃ��݌ɂ̃��P�[�V�������擾
Call FetchExcellJanInventory

Application.ScreenUpdating = True

Worksheets("��z���ʓ��̓V�[�g").Activate

Call CheckNonArrival

'�f�[�^�o�͂̃{�^����z�u
With Worksheets("��z���ʓ��̓V�[�g")

    Dim EndRow As Long
    EndRow = Worksheets("��z���ʓ��̓V�[�g").UsedRange.Rows.Count

    With .Buttons.Add( _
        Range("B" & EndRow).Left - 20, _
        Range("B" & EndRow).Top + 20, _
        200, _
        30 _
        )
        
        .OnAction = "BuildPurcahseData"
        .Characters.Text = "�����V�X�e���p�f�[�^�o��"
        .Name = "BuidDataButton"
        
    End With

    .Range("A2").Activate

End With

ActiveWindow.ScrollColumn = 1
ActiveWindow.ScrollRow = 1

MsgBox Prompt:="��z���ʓ��̓V�[�g�A�f�[�^���͊���" & vbLf & "�ۗ��`�F�b�N�A��z���ʂ̏C�����s���Ă��������B", Buttons:=vbInformation

End Sub

Private Sub FetchSyokonData()
'�������i�}�X�^�[�̃f�[�^�擾

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
        If Len(Code) > 6 Then
            r.NumberFormatLocal = "@"
            r.Value = IIf(Len(DbRs("���i�R�[�h")) = 5, "0" & DbRs("���i�R�[�h"), DbRs("���i�R�[�h"))
        End If
    
    End If

Next

End Sub

Private Sub FetchExcellForPurchase()
'�����p���i���u�b�N���d����E���b�g�E��z�����l�̎擾

Dim DataBook As Workbook, DataSheet As Worksheet, PurDataCodeRange As Range, PurDataJanRange As Range

Set DataSheet = FetchWorkBook("\\Server02\���i��\�l�b�g�̔��֘A\�����֘A\�����p���i���.xlsm").Worksheets("���i���")

DataSheet.Activate

Set PurDataJanRange = DataSheet.Range(Cells(1, 1), Cells(DataSheet.UsedRange.Rows.Count, 1))
Set PurDataCodeRange = DataSheet.Range(Cells(1, 2), Cells(DataSheet.UsedRange.Rows.Count, 2))

ThisWorkbook.Activate
Dim CodeRange As Range, r As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange

    Dim Code As String, HitRow As Double
        
    Code = r.Value

    On Error Resume Next
        
        '�G���[���ɑO��i�[�����l�̂܂܂ɂȂ�̂Ŗ����I�ɏ�����
        HitRow = 0
        
        HitRow = WorksheetFunction.Match(Code, PurDataCodeRange, 0)

        If Err Then
            Err.Clear
            HitRow = WorksheetFunction.Match(Code, PurDataJanRange, 0)
            
            If Err And IsEmpty(Cells(r.Row, 4).Value) Then '�d����R�[�h�������ɂ��G�N�Z���ɂ��Ȃ���΁A�����ł��Ȃ��̂Œ��ӏ���������
                Cells(r.Row, 2).Value = "�����p���i��� �Y��JAN�Ȃ�"
            End If
        
        End If
            
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

    On Error GoTo 0

Next

End Sub

Private Sub FetchExcellJanInventory()
'�I�Ȃ��݌ɕ\�f�[�^�̊m�F

Const NOLOCATION_INVENTRY_EXCELL As String = "\\server02\���i��\�l�b�g�̔��֘A\�I���݌Ɋm�F�\.xlsm"
Const INVENTRY_SHEET As String = "�I���f�[�^"

Workbooks.Open FileName:=NOLOCATION_INVENTRY_EXCELL, ReadOnly:=True

With ThisWorkbook.Worksheets("��z���ʓ��̓V�[�g")
    Dim CodeRange As Range, r As Range
    Set CodeRange = .Range(.Cells(2, 7), .Cells(2, 7).End(xlDown))
End With

With Workbooks(Dir(NOLOCATION_INVENTRY_EXCELL)).Worksheets(INVENTRY_SHEET)
    Dim InventryRange As Range
    Set InventryRange = .Range(.Cells(1, 2), .Cells(1, 2).End(xlDown))
End With

For Each r In CodeRange

    Dim Code As String, HitRow As Double, Location As String, StockQuantity As Long
    
    Code = r.Value

    On Error Resume Next
    
    '�G���[���ɑO��i�[�����l�̂܂܂ɂȂ�̂Ŗ����I�ɏ�����
    HitRow = 0
    
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

Workbooks(Dir(NOLOCATION_INVENTRY_EXCELL)).Close SaveChanges:=False

End Sub

Private Sub CalcPurchaseQuantity()
'��z�˗����ʂ���A���b�g�P�ʁE�����P�ʂŊۂ߂����������ʂ��Z�o���AA��֓����B

Dim CodeRange As Range, r As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange
    Dim i As Long
    i = r.Row
    
    Dim Rot As Double, Qty As Double, RequestQty As Double
    Rot = CDbl(Cells(i, 12).Value)
    
    If IsEmpty(Rot) Or Rot = 0 Then
        Rot = 1
    End If
    
    RequestQty = CDbl(Cells(i, 9).Value)
    
    '�Z�C�����O�֐��ɂă��b�g���̔{���Ŏ�z�˗������ۂ߂�
    Qty = WorksheetFunction.Ceiling(RequestQty, Rot)

    Cells(i, 1).Value = Qty

    '���b�g��1�łȂ��ꍇ�́A��z���ʂ��C������邽�ߋ����\��
    If Rot <> 1 Then
    
        With Union(Cells(i, 1), Cells(i, 9)).Interior
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
    
    End If
    
Next

End Sub

Private Function GetKubunLabel(ByVal KubunCode As Variant) As String
'���i�}�X�^�ł͋敪��1�`9�̐����Ȃ̂ŁA�\�����Œu��������B
'����-�敪���̑g�͖��̃}�X�^�Ɋi�[����Ă���B

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

Private Sub FetchPickupFlag()
'����̎d����ɂ��āA�d���惊�X�g�̃V�[�g����Vlookup�֐��ɂĈ���̋敪�ԍ����擾����B
'����ȊO�͔����敪�� 2

'�d����R�[�h�̃����W���Z�b�g�A1�Z������Vlookup���s��
Dim CodeRange As Range, r As Range, VendorsRange As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown)).Offset(0, -3)

'�d����R�[�h��Vlookup���ĒT�������W
Set VendorsRange = ThisWorkbook.Worksheets("�d���惊�X�g").Range("A1").CurrentRegion

For Each r In CodeRange

    Dim VendorCode As String, DeliveryDiv As Integer
    
    VendorCode = r.Value
    
    On Error Resume Next
    
        DeliveryDiv = WorksheetFunction.VLookup(VendorCode, VendorsRange, 3, False)
        
        If Err Or DeliveryDiv = 0 Then
            DeliveryDiv = 2
        End If
        
    On Error GoTo 0
    
    Cells(r.Row, 11).Value = DeliveryDiv
    
Next

End Sub

