VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OpPanel 
   Caption         =   "���t�[��������"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6405
   OleObjectBlob   =   "OpPanel.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "OpPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�S�v���V�[�W���ŋ��ʂ���t�@�C���̏��݂ƃt�@�C����

'���ׂƒ����w�b�_�[�̍ݏ�
Const MEISAI_PATH As String = "\\MOS10\Users\mos10\Desktop\���t�[\meisai.csv"
Const TYUMON_PATH As String = "\\MOS10\Users\mos10\Desktop\���t�[\tyumon_H.csv"

Private Sub CommandButton1_Click()
    
    ����󒍃t�@�C���Ǎ�
    
End Sub

Private Sub CommandButton12_Click()
    
    CheckPickingProducts
    
End Sub

Private Sub CommandButton13_Click()
    
    FetchFaxReply.FetchFaxReply

End Sub

Private Sub CommandButton14_Click()
    
    �����ԍ��ꊇ�t�@�C���쐬
    
End Sub

Private Sub CommandButton15_Click()
    
    �x���`�F�b�N
    
End Sub

Private Sub CommandButton16_Click()
    '�A�E�g���C���̓W�J�A�܂肽���݂ɂ���  http://www.relief.jp/itnote/archives/017927.php
    
    OrderSheet.Outline.ShowLevels ColumnLevels:=2
    
End Sub

Private Sub CommandButton17_Click()

    OrderSheet.setProtect

End Sub

Private Sub CommandButton18_Click()
    
    OrderSheet.setUnprotect
    MsgBox "30�b�Ԃ̂݁A���c�ꗗ�̋L��/�폜���t���[�ɂȂ�܂��B"
    End

End Sub

Private Sub CommandButton19_Click()

    Call checkMesaiFileExistance

End Sub

Private Sub CommandButton2_Click()
    
    Meisai�ʓǍ�

End Sub
Private Sub CommandButton3_Click()

    tyumon_H�ʓǍ�

End Sub
Private Sub CommandButton4_Click()

    fetchShippingDone

End Sub
Private Sub CommandButton5_Click()

    Unload Me

End Sub

Private Sub CommandButton8_Click()
    
    archiveCompletedOrder

End Sub

Private Sub UserForm_Initialize()
'���[�U�[�t�H�[�����J�������̏���

'�y�[�W1��\��
Me.MultiPage1.Value = 0

Dim PickingFilePath As String

TextBox1.Value = "���݃`�F�b�N ��"
TextBox2.Value = "���݃`�F�b�N ��"


If Format(Now(), "hh") < 10 Then '10���܂ł̃t�H�[���I�[�v���Ȃ�AMeisai�t�@�C�����݃`�F�b�N

    Call checkMesaiFileExistance

Else

    TextBox1.Value = "�v �蓮�`�F�b�N"
    TextBox2.Value = "�v �蓮�`�F�b�N"

End If

PickingFilePath = Range("PickingSheetFolder").Value

If Right(PickingFilePath, 1) <> "\" Then PickingFilePath = PickingFilePath & "\" '����\�}�[�N�łȂ��Ƃ��⊮

PickingFileName = Range("PickingSheetBaseName").Value


'�e�L�X�gBox4�A�s�b�L���O�t�@�C���̃t�@�C�����
TextBox4.Value = Dir(PickingFilePath & PickingFileName & "*")

'�e�L�X�gBox5�A�o�׈ꗗ�ڍׂ̃t�@�C�����

If Format(Now(), "hh") > 15 Then '15���߂��Ẵt�H�[���I�[�v���Ȃ�A�o�׈ꗗ�ڍׂ̃t�@�C�����݃`�F�b�N
    
'�o�׈ꗗ�ڍׂ̂��肩 �t�@�C���`�F�b�N�Ɏg���Ă܂�
    Const REPORT_FILE_PATH As String = "\\server02\���i��\�l�b�g�̔��֘A\�o�גʒm\�o�גʒm_�y�V\���o�׈ꗗ�ڍ�\"
    Const REPORT_FILE_BASE As String = "�o�׈ꗗ�ڍ�_"
    
    Dim ReportFileName As String
    ReportFileName = REPORT_FILE_BASE & Format(Date, "yymmdd") & ".xlsx"
    TextBox5.Value = Dir(REPORT_FILE_PATH & ReportFileName)
    
End If

End Sub

Private Sub checkMesaiFileExistance()

'�t�@�C������I�u�W�F�N�g����
Dim FSO As Object
Set FSO = New FileSystemObject

If FSO.FileExists(MEISAI_PATH) Then
    
    TextBox1.Value = FSO.GetFile(MEISAI_PATH).DateCreated
    TextBox2.Locked = True
    
Else
    
    TextBox1.Value = "�t�@�C���Ȃ�"

End If

If FSO.FileExists(TYUMON_PATH) Then
    
    TextBox2.Value = FSO.GetFile(TYUMON_PATH).DateCreated
    TextBox2.Locked = True
    
Else
    
    TextBox2.Value = "�t�@�C���Ȃ�"

End If

End Sub

Private Sub fetchShippingDone()

Dim LineBuf As Variant
Dim order As Variant

'�A���󋵃V�[�g��OrderSheet�̒����ԍ��̃����W
Dim search_range As Range
Set search_range = OrderSheet.Cells(2, 2).Resize(OrderSheet.Cells(2, 2).SpecialCells(xlCellTypeLastCell).Row, 1)

'���[�v���Ŏg��Find�֌W�̃����W
Dim firstCell As Range
Dim FoundCell As Range

' �t�@�C���_�C�A���O����p�X���w�肵�āAFSO�ŊJ��
Dim file_path As String
file_path = fetchOrderCsv.setCsvPath("shipping.csv")

If file_path = "" Then
    MsgBox "�t�@�C���w�肪�L�����Z������܂����B"
    Exit Sub
End If

Dim FSO As Object
Set FSO = New FileSystemObject

' CSV���e�L�X�g�X�g���[���Ƃ��ď�������
Dim TS As Textstream
Set TS = FSO.OpenTextFile(file_path, ForReading)
       
Do Until TS.AtEndOfStream
    
'�����ԍ��A����掁���A�i�����j�󋵁A�₢���킹�ԍ���z��tmp�ɓ����

    LineBuf = Split(TS.ReadLine, ",")
    
    'tmp[0]=OrderID=Column"A"
    'tmp[1]=Ship name=����於=Column"B",
    'tmp[2]=status=��=Column"C"
    'tmp[3]=shipDate=������=Column"D"
    'tmp[4]=Shipping Number=�₢���킹�ԍ�=Column"E"
    
    tmp = Array(LineBuf(0), LineBuf(1), LineBuf(2), LineBuf(3), LineBuf(4))
    
    For j = 0 To UBound(tmp)
        tmp(j) = Trim(Replace(tmp(j), Chr(34), "")) 'chr(34)�� " [���p��d���p��]
    
    Next

'���c�ꗗ�V�[�g�̊Y�����钍���ԍ��ɓǂݎ������������
                
    Set FoundCell = search_range.Find(what:=tmp(0))
    
    If Not FoundCell Is Nothing Then
        
        FirstCellAddress = FoundCell.Address
        
        If Not OrderSheet.Cells(FoundCell.Row, 3).Value = tmp(1) Then '�����Җ����Ȃ���΍ŏ���FoundCell�ɓ����
        
            OrderSheet.Cells(FoundCell.Row, 3) = tmp(1)
        
        End If
                
        Do
            '�����󋵁u�ρv��Find���Č������������ԍ����ׂĂɓ����
            'IF �󋵂����� AND �����Z������ AND �₢���킹�ԍ������� Then �����󋵗�ɍρA��������ɔ�����
            'TODO:FIND�Ń��[�v�񂳂Ȃ��悤�Ƀ��t�@�N�^�����O
            
            If Cells(FoundCell.Row, "O").Value = "" And Not tmp(4) = "" Then
                OrderSheet.Cells(FoundCell.Row, 15) = "��" '������=������O��
                OrderSheet.Cells(FoundCell.Row, 16) = tmp(3)
    
            End If
            
            Set FoundCell = Cells.FindNext(FoundCell)
            
            If FoundCell Is Nothing Or FoundCell.Address = FirstCellAddress Then Exit Do
        
         Loop

    End If

Loop


' �w��t�@�C����CLOSE
TS.Close
Set TS = Nothing
Set FSO = Nothing

OpPanel.Hide

Call CheckBelate.�x���`�F�b�N

ThisWorkbook.Save

End Sub

Private Sub archiveCompletedOrder()

'�O�X���ŏI���܂ł̔����ς݁A�L�����Z����ʃV�[�g�Ɉړ����܂�

Application.ScreenUpdating = False

'�I�[�g�t�B���^�[���������Ă���Έ�U�������܂��A�R�s�[��̍s����\���ɂȂ邽��
'sheet�I�u�W�F�N�g�ɁA�I�[�g�t�B���^�[�̗L��������FilterMode�v���p�e�B:Boolean�^������܂�
If OrderSheet.FilterMode Then OrderSheet.Range("A1").CurrentRegion.AutoFilter

'�{�����t���獡���́u���v���������O���ŏI�����Z�o
Dim LastDay As Date
LastDay = DateAdd("d", -(Format(Date, "d")), Date)

Dim i As Long
i = 2

With ThisWorkbook.Sheets("���c�ꗗ")

'Cell(1,1)�̓��t���r���đO���ŏI���𒴂��Ȃ����菈��������

Do Until DateDiff("d", LastDay, .Cells(i, 1)) > 1
    
    '�󗓂̏ꍇ�������L���X�g�Ŕ�r���Ă��܂��A�������~�܂�Ȃ��ꍇ�����肦��
    If IsEmpty(.Cells(i, 1)) Then Exit Sub
    
    If .Cells(i, "O").Value = "��" Then

        .Rows(i).Cut Destination:=Sheets("����").Rows(Sheets("����").UsedRange.Rows.Count + 1)
        .Rows(i).Delete
        
    ElseIf OrderSheet.Cells(i, "O").Value = "�L�����Z��" Then
    
        .Rows(i).Cut Destination:=Sheets("�L�����Z��").Rows(Sheets("�L�����Z��").UsedRange.Rows.Count + 1)
        .Rows(i).Delete
    
    Else
        
        '��O����or�L�����Z���łȂ��ꍇ���󗓂̎��A�s�|�C���^��i�߂�
        i = i + 1
    
    End If

Loop

End With

Application.ScreenUpdating = True

OrderSheet.Activate

'�{�^�����Ĕz�u
OrderSheet.Shapes("ShowFormButton").Top = OrderSheet.Range("A1").End(xlDown).Offset(2, 1).Top
OrderSheet.Shapes("ButtonHideWish").Top = OrderSheet.Range("A1").End(xlDown).Offset(2, 1).Top

End Sub

