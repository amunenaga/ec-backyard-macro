Attribute VB_Name = "FetchPickingSheet"
Option Explicit

'���̃��W���[������J�����s�b�L���O�V�[�g�t�@�C�����N���[�Y����̂ɕK�v�ȃp�u���b�N�ϐ�
'
Public IsFileNewOpen As Boolean
Public PickingFileName As String

Function CheckPickingProducts(Optional IsMsgBox As Boolean = True)

'�}�N���̈ꗗ�ɏo�������Ȃ��̂�Function��`�ɂ��Ă��܂��B
'���t�[�X�J�C�v�����t�t�@�C���ŁA�΂ɓh���Ă��Ȃ�=�s�b�L���O�ł��Ȃ��������i���A
'���c�V�[�g�A�Z���^�[�݌ɗ�Ɂu�Ȃ��v�]�L�A�{�����t�Ŏ�z���������Ƃ݂Ȃ��ē��t����܂��B

'�{����������荞��ł��邩�`�F�b�N
Application.ScreenUpdating = False
Application.DisplayAlerts = False

If LogSheet.Range("LastFetchNewOrder") <> Date Then
    
    Call fetchOrderCsv.����󒍃t�@�C���Ǎ�
    LogSheet.Range("B7").Value = Date

End If

Application.DisplayAlerts = True

'�u���t�[�X�J�C�v��xx.xlsx�v���J�����������̃`�F�b�N�ς݂̏��i���X�g�̃G�N�Z���t�@�C���J��
'PickingFileName = Range("PickingSheetBaseName") & Format(Date, "mmdd") & ".xlsx"

'�t�H�[����TextBox4��
PickingFileName = OpPanel.TextBox4


Dim PickingFilePath As String

PickingFilePath = Range("PickingSheetFolder").Value

If Right(PickingFilePath, 1) <> "\" Then PickingFilePath = PickingFilePath & "\" '����\�}�[�N�łȂ��Ƃ��⊮

PickingFilePath = PickingFilePath & PickingFileName


''�u�t�@�C�����J���v�̃t�H�[���Ńt�@�C�����w�� �ꉞ�c���܂�
'PickingFilePath = Application.GetOpenFilename("�G�N�Z���t�@�C��,*.xls?", , "���t�[�s�b�L���O���X�g���w��")

Dim WsPicking As Worksheet
Dim wb As Workbook

'�s�b�L���O�V�[�g���J���Ă���΂��̂܂ܗ��p����A
'Todo:��^�����Ő؂�o���Ă��܂������������AInvoices�N���X�ɂ������悤�ȏ���������
For Each wb In Workbooks
    If wb.Name = PickingFileName Then
        Set WsPicking = wb.Sheets(1)
    End If
Next wb

'���[�N�u�b�N���J���ăZ�b�g
If WsPicking Is Nothing Then

    '�l�b�g�̔��֘A�̏���̃t�H���_�Ƀs�b�L���O�V�[�g���Ȃ��ꍇ�AExit
    If Not Dir(PickingFilePath) Like "*.xlsx" Then
        
        MsgBox "�s�b�L���O�V�[�g�̓]�L���ł��܂���ł����B" & vbLf & _
                "�s�b�L���O�V�[�g�t�@�C���Ȃ��B���t�[�s�b�L���O�V�[�g�̃t�@�C���L���A�t�@�C�������m�F" & vbLf & _
                "���̏����͌p�����ĉ\�ł��B"

        Exit Function
    
    End If

    Set wb = Workbooks.Open(PickingFilePath)
    Set WsPicking = wb.Sheets(1)
    
    IsFileNewOpen = True
        
End If

'�s�b�L���O�Ώۂ̏��i�R�[�h�����W
Dim MaxRow As Integer
MaxRow = WsPicking.Range("B1").SpecialCells(xlCellTypeLastCell).Row

'----�s�b�L���O�V�[�g�̃I�[�v����������-----

Dim TodaysOrders As Dictionary
Set TodaysOrders = OrderSheet.getTodaysOrders '�{���󒍂�OrderList�𐶐�


'��U�A�S�ẴA�C�e����IsPickingDone��False�ɃZ�b�g�AOrderObject�����̓x�ɃZ����ǂ�ŋ�Ȃ�True�Ȃ̂�"
Dim v As Variant
Dim w As Variant
For Each v In TodaysOrders
    For Each w In TodaysOrders(v).Products
        TodaysOrders(v).Products(w).IsPickingDone = False
    Next
Next

'�s�b�L���O�V�[�g���璍���Җ��E���i�R�[�h���擾���āATodaysOrder�Ɠ˂����킹��
Dim i As Integer
For i = 2 To MaxRow
    
    Dim CurrentBuyerName As String
    CurrentBuyerName = WsPicking.Cells(i, 1).Value
    
    Dim CurrentCode As String
    CurrentCode = WsPicking.Cells(i, 2).Value
        
    Dim CurrentNote As String
    CurrentNote = WsPicking.Cells(i, 8).Value
        
    '�R�[�h�����t�[�̌`���ɕϊ� 012345->12345
    If CurrentCode Like "0#####" Then CurrentCode = Right(CurrentCode, 5)
    
    Dim o As order
    
    '�w�i�F�����ł͂Ȃ����i=�Z���^�[�݌ɗL��A�s�b�L���O�\
    If Not WsPicking.Cells(i, 1).Interior.Color = 16777215 Then
        
        '�����Җ�����A�ǂ̒����̏��i������
        Set o = FindByBuyerName(CurrentBuyerName, CurrentCode, TodaysOrders)
        
        '���̒����̏��i�I�u�W�F�N�g�Ƀs�b�L���O�̃t���O��o�^
        o.Products(CurrentCode).IsPickingDone = True
    
    End If
    
    'H��ɉ��������Ă遁����Ŕc�����Ă���݌ɏ�
    If WsPicking.Cells(i, 8).Value <> "" Then
                '�����Җ�����A�ǂ̒����̏��i������
                
        Set o = FindByBuyerName(CurrentBuyerName, CurrentCode, TodaysOrders)
        
        '���̒����̏��i�I�u�W�F�N�g�Ƀs�b�L���O�̃t���O��o�^
        o.Products(CurrentCode).CenterStockState = WsPicking.Cells(i, 8).Value

    End If
    
Next i

'TodaysOrder�̊e�����̊e���i��IsPickingDone���`�F�b�N����


'Dim w As Variant
'Dim v As Variant

For Each v In TodaysOrders
    For Each w In TodaysOrders(v).Products
        If TodaysOrders(v).Products(w).IsPickingDone = False Then
            
            '�s�b�L���O�X�e�[�^�X���V�[�g�ɓ]�L�AFalse���Ɓu�Ȃ��v�{�{����z����
            'Todo:Order��Product�I�u�W�F�N�g��n��
            Call OrderSheet.writePickingStatus(CStr(v), CStr(w), TodaysOrders(v).Products(w).CenterStockState)
        
        End If
    
    Next
Next

'�`�F�b�N����LogSheet�ɏ�������
LogSheet.Range("LastUpdatePickingSheet") = Date

ThisWorkbook.Save

Application.ScreenUpdating = False

If IsMsgBox Then
    
    MsgBox prompt:="�s�b�L���O�t�@�C���̓]�L����", Buttons:=vbInformation

End If

End Function

Private Function FindByBuyerName(Name As String, Code As String, OrderList As Dictionary) As order
'�������X�g�z��ƒ����Җ����󂯎���āAOrder�I�u�W�F�N�g��Ԃ��B
'�����Җ��Œ�����T���āA���̎󒍃A�C�e��Products�ɊY���R�[�h�����邩����A
'Order�I�u�W�F�N�g��Ԃ�

'Order�̔z����܂����O�Œ��ׂāA�Y�������Products���̃R�[�h�𒲂ׂ�
Dim v As Variant
For Each v In OrderList

    If OrderList(v).BuyerName = Name Then
        
        Dim w As Variant
        
        For Each w In OrderList(v).Products
            
            If OrderList(v).Products.Exists(Code) Then
            
                Set FindByBuyerName = OrderList(v)
            
                Exit Function
            
            End If
        
        Next w
    
    End If

Next v

End Function
