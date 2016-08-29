Attribute VB_Name = "fetchOrderCsv"
Option Explicit

'���ׂƒ����w�b�_�[�̂���t�H���_���w��A�Ō�K��\�}�[�N
Const CSV_PATH As String = "\\MOS10\Users\mos10\Desktop\���t�[\"

'���������J�E���^
Dim OrderCount As Long

Sub ����󒍃t�@�C���Ǎ�()

Dim LineBuf As Variant
Dim OrderDetail As Variant

'�t�@�C������I�u�W�F�N�g����
Dim FSO As New FileSystemObject

' Meisai.csv��tyumon_H.csv�̑��݃`�F�b�N
Dim MeisaiPath As String
MeisaiPath = CSV_PATH & "Meisai.csv"

If FSO.FileExists(MeisaiPath) = False Then
    
    MsgBox "Meisai.csv��������܂���" & vbLf & _
            CreateObject("WScript.Network").UserName & "�ł�MOS10�̃f�[�^���Q�Ƃł��Ȃ��̂ŁA��PC�Ŏ��s���Ă��������B" & vbLf & _
            "�������́A���t�[�̊Ǘ���ʂ���_�E�����[�h���āA�u�ʓǍ��v�Ŏw�肵�Ă��������B" & vbLf & _
            vbLf & "�������I�����܂��B"
    
    End

End If

Dim TyumonhPath As String
TyumonhPath = CSV_PATH & "tyumon_H.csv"

If FSO.FileExists(TyumonhPath) = False Then
    
    MsgBox "tyumon_H.csv��������܂���" & vbLf & _
            CreateObject("WScript.Network").UserName & "�ł�MOS10�̃f�[�^���Q�Ƃł��Ȃ��̂ŁA��PC�Ŏ��s���Ă��������B" & vbLf & _
            "�������́A���t�[�̊Ǘ���ʂ���_�E�����[�h���āA�u�ʓǍ��v�Ŏw�肵�Ă��������B" & vbLf & _
            vbLf & "�������I�����܂��B"
    
    End

End If

' �{�����A�Ǎ��ς��m�F
If LogSheet.Range("LastFetchNewOrder").Value = Date Then
    
    Dim mb As Variant
    mb = MsgBox("�{�����͓Ǎ��ςł��B" & vbLf & "�����𑱂��܂����H", vbYesNo + vbExclamation)
        
    If mb = vbNo Then
        MsgBox "�������L�����Z�����܂����B"
        Exit Sub
    
    End If
End If

Call readMeisai(MeisaiPath)

Call sortOrderId

Call readTyumonH(TyumonhPath)

LogSheet.Range("LastFetchNewOrder").Value = Date

ThisWorkbook.Save

'�v�]���W�J���܂��B
OrderSheet.Outline.ShowLevels ColumnLevels:=2

MsgBox Prompt:=Format(Date, "m��d��") & " �󒍕�  " & OrderCount & "��" & vbLf & " �Ǎ��������܂����B" _
    , Buttons:=vbInformation

End Sub

Function Meisai�ʓǍ�(Optional str As String = "") As Variant

Dim FilePath As String

'meisai.csv���t�@�C���_�C�A���O�Ŏw��"
FilePath = setCsvPath("meisai.csv")

If FilePath = "" Then
    MsgBox "�t�@�C���w�肪�L�����Z������܂����B"
    Exit Function
End If

Call readMeisai(FilePath)

Call sortOrderId

MsgBox "�Ǎ�����"

End Function

Function tyumon_H�ʓǍ�(Optional str As String = "") As Variant

Dim FilePath As String

'tyumon_H.csv���t�@�C���_�C�A���O����w�肷��
FilePath = setCsvPath("tyumon_H.csv")

If FilePath = "" Then
    MsgBox "�t�@�C���w�肪�L�����Z������܂����B"
    Exit Function
End If

Call readTyumonH(FilePath)

MsgBox "�Ǎ�����"

End Function

Function setCsvPath(CsvName As String)

'Application�I�u�W�F�N�g�擾
Dim xlApp As Application
Set xlApp = Application

'��t�@�C�����J����̃t�H�[���Ńt�@�C�����̎w����󂯂�
Dim FileName As Variant
FileName = xlApp.GetOpenFilename(filefilter:="CSV�t�@�C��,*.csv" _
                                    , Title:=CsvName & "���w�肵�Ă�������")

'�L�����Z�����ꂽ�ꍇ��False���Ԃ�̂ňȍ~�̏����͍s�Ȃ�Ȃ�
If VarType(FileName) = vbBoolean Then End

setCsvPath = FileName
    
End Function

Private Sub readMeisai(Path As String)

'Meisai.CSV��OrderSheet=�����ꗗ�ɒǋL����

'�_�u��`�F�b�N�̂��߂ɓǍ��O�̒��c�V�[�g�̃����W���w��
Dim LastRow As Integer
LastRow = OrderSheet.Cells.SpecialCells(xlCellTypeLastCell).Row

Dim RngLoadedOrders As Range
Set RngLoadedOrders = OrderSheet.Range(Cells(2, 2), Cells(LastRow, 2))

Dim FSO As Object
Set FSO = New FileSystemObject

Dim TS As Textstream
Set TS = FSO.OpenTextFile(Path, ForReading)


Dim i As Long
i = LastRow '���ڍs�͏o�͂��Ȃ��̂ŁAi�͏I�[�s����J�n
    
Dim OrderCount As Long
OrderCount = 0
    
Do Until TS.AtEndOfStream
    
' �s�����o���ĕK�v�ȍ��ڂ݂̂�z��ɓ��꒼��
    Dim LineBuf As Variant
    LineBuf = Split(TS.ReadLine, """,""")
    
    Dim OrderDetail As Variant
    
    '0=1���=�����ԍ��A1=2���=1�������ŉ��A�C�e���ڂ��A2=3���=���A4=5���=�R�[�h 5=6���=���i��
    '�n�[�h�R�[�f�B���O���Ă���̂ŁA�����Ǘ���ʂ���o�͍��ڂ�ύX������A�ǂݎ��Ȃ��Ȃ�܂��B

    OrderDetail = Array(LineBuf(0), LineBuf(1), LineBuf(2), LineBuf(3), LineBuf(4))
    
    Dim j As Integer
    For j = 0 To UBound(OrderDetail)
        OrderDetail(j) = Trim(Replace(OrderDetail(j), Chr(34), "")) 'chr(34)�� " [���p��d���p��]�炵��
    
    Next
    
    '���[�v���ڂł̓w�b�_�[�Ȃ̂ŁA�C���N�������g�֔��
    If OrderDetail(0) = "Order ID" Then GoTo increment
    
    '�����ԍ������ɓǍ��ς̃Z���͈͂ɂ���ꍇ���C���N�������g��
    If WorksheetFunction.CountIf(RngLoadedOrders, OrderDetail(0)) <= 0 Then
    
        Cells(i, 1).Value = Date
        Cells(i, 2).Value = OrderDetail(0)
        Cells(i, 4).Value = OrderDetail(1)
        Cells(i, 5).Value = OrderDetail(3)
        Cells(i, 6).Value = OrderDetail(4)
        Cells(i, 7).Value = OrderDetail(2)
    
    Else
        GoTo increment
    
    End If
    
increment:
    i = i + 1

Loop

Call sortOrderId

'���[�U�[�t�H�[���Ăяo���{�^���̈ʒu����
OrderSheet.Shapes("ShowFormButton").Top = OrderSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Offset(2, 1).Top
'OrderSheet.Shapes("hideWishCol").Top = OrderSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Offset(2, 1).Top

' TextStream��ؒf
TS.Close

End Sub

Private Sub readTyumonH(Path As String)

Dim FSO As Object
Set FSO = New FileSystemObject

Dim TS As Textstream
'Set TS = FSO.OpenTextFile("C:\Users\mos9\Downloads\tyumon_H_3.csv", ForReading)
Set TS = FSO.OpenTextFile(Path, ForReading)

'�Ǎ��ϒ����ԍ��̃����W���Z�b�g
Dim LoadedOrderRange As Range
Set LoadedOrderRange = OrderSheet.Cells(2, 2).Resize(OrderSheet.Cells(2, 2).SpecialCells(xlCellTypeLastCell).Row, 1)

Do Until TS.AtEndOfStream
    
' �s�����o���ĕK�v�ȍ��ڂ݂̂�z��ɓ��꒼��
    Dim LineBuf As Variant
    LineBuf = Split(TS.ReadLine, ",")
    
    '0=1���=�����ԍ��A�����Җ��A�v�]�A���ϕ��@�A�N�[�|���l����
    Dim order As Variant
    order = Array(LineBuf(0), LineBuf(5), LineBuf(36), LineBuf(34), LineBuf(43))
        
    Dim j As Integer
    For j = 0 To UBound(order)
        order(j) = Trim(Replace(order(j), Chr(34), "")) 'chr(34)�� " [���p��d���p��]
    
    Next

    '�����ԍ��̍s�𒲂ׂ�
    Dim FindRow As Double 'Match�֐��̕Ԓl��Double�^
    
    On Error Resume Next
        
        FindRow = WorksheetFunction.Match(CDbl(order(0)), LoadedOrderRange, 0) + 1  '�R�[�h�����W��B2����n�܂�̂ōs����1��������
        
        If Err Then GoTo continue
        
    On Error GoTo 0
    
    Range("C" & FindRow).Value = order(1) '�����Җ�������
    
    '��U�Atmp�ɔ��l�����e��ێ�
    Dim tmp As String
    tmp = Range("S" & FindRow).Value
    
    '�N�[�|�����p��������E��s�U���E���t�[�}�l�[���� �m�F���Ĕ��l���֒ǋL
    If order(3) = "payment_d1" And order(4) < 0 Then tmp = "����� �N�[�|�����p "
    If order(3) = "payment_b1" Then tmp = tmp & "�U�� �����ē� ��"
    If order(3) = "payment_a16" Then tmp = tmp & "Yahoo!�}�l�[����"
    
    Range("S" & FindRow).Value = tmp 'tmp���Z���ɏ����߂�
    
    If order(2) <> "" Then Range("Q" & FindRow).Value = order(2) '�v�]��]�L
    
        
    OrderCount = OrderCount + 1
    
continue:
    
Loop

' �I�u�W�F�N�g��j��
TS.Close
Set TS = Nothing
Set FSO = Nothing

End Sub

Sub �����X�e�[�^�XCSV�Ǎ�()

Dim LineBuf As Variant
Dim order As Variant

'�A���󋵃V�[�g��OrderSheet�̒����ԍ��̃����W
Dim IdRange As Range
Set IdRange = OrderSheet.Cells(2, 2).Resize(OrderSheet.Cells(2, 2).SpecialCells(xlCellTypeLastCell).Row, 1)

'���[�v���Ŏg��Find�֌W�̃����W
Dim firstCell As Range
Dim FoundCell As Range

' �t�@�C���_�C�A���O����p�X���w�肵�āAFSO�ŊJ��
Dim Path As String
Path = fetchOrderCsv.setCsvPath("order_process_status.csv")

If Path = "" Then
    MsgBox "�t�@�C���w�肪�L�����Z������܂����B"
    Exit Sub

End If

Dim FSO As Object
Set FSO = New FileSystemObject

' CSV���e�L�X�g�X�g���[���Ƃ��ď�������
Dim TS As Textstream
Set TS = FSO.OpenTextFile(Path, ForReading)
       
'�w�b�_�[���`�F�b�N
LineBuf = Split(TS.ReadLine, ",")

If Trim(Replace(LineBuf(1), Chr(34), "")) <> "OrderStatus" Then
    MsgBox "CSV�t�@�C���������X�e�[�^�X�ꗗ�ł͂���܂���B�����𒆎~���܂�"
    Exit Sub
End If

'
Call �S�Ă̔����󋵂�\��
 
Do Until TS.AtEndOfStream
    
    '�����ԍ��A����掁���A�i�����j�󋵁A�₢���킹�ԍ���z��tmp�ɓ����

    LineBuf = Split(TS.ReadLine, ",")
    
    'tmp[0]=OrderID=Column"B"
    'tmp[1]=
    
    Dim tmp As Variant
    tmp = Array(LineBuf(0), LineBuf(1), LineBuf(2), LineBuf(3))
    
    Dim j As Long
    For j = 0 To UBound(tmp)
        tmp(j) = Trim(Replace(tmp(j), Chr(34), "")) 'chr(34)�� " [���p��d���p��]
    
    Next

    '���c�ꗗ�V�[�g�̊Y�����钍���ԍ��ɓǂݎ������������
                
    Set FoundCell = IdRange.Find(what:=tmp(0))
    
    If Not FoundCell Is Nothing Then
           
        Dim FirstCellAddress As String
        FirstCellAddress = FoundCell.Address
        

                       
        Do
            '�����󋵂�Find���Č������������ԍ����ׂĂɓ����A�㏑���ł悢�B
            
            OrderSheet.Cells(FoundCell.Row, 18) = tmp(1)
            
            Set FoundCell = Cells.FindNext(FoundCell)
            
            If FoundCell Is Nothing Or FoundCell.Address = FirstCellAddress Then Exit Do
        
         Loop

    End If

Loop


' �w��t�@�C����CLOSE
TS.Close
Set TS = Nothing
Set FSO = Nothing

'�������̂ݕ\���ɕύX
Call �������̂ݕ\��

OpPanel.Hide

ThisWorkbook.Save

End

End Sub


Private Sub sortOrderId()

'OrderID�̗��T�� B2���ߑł��ł��悭�Ȃ����ȁH
Dim col_orderID As Range
Set col_orderID = OrderSheet.Rows(1).Find("Order ID")

With OrderSheet.Sort

    .SortFields.Clear '��U�A�������N���A�[����
    .SortFields.Add key:=col_orderID, order:=xlAscending '�\�[�g�������Z�b�g

    '�\�[�g�����s����
    .SetRange Range("A1").CurrentRegion
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply '�\�[�g�K�p

End With

End Sub
