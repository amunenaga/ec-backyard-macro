Attribute VB_Name = "FetchPickingSheet"
Option Explicit

'�����ԍ��A�����Җ��ȂǁA���o�����Ɨ�ԍ����i�[����Head�^���`
'�쐬�҂�쐬���ɂ���āA��̈ʒu���ς�邱�Ƃ�����̂ŁA������肷��
'��{�I�ɂ́A�}�j���A���ƃe���v���[�g�Ńf�[�^�����Ƃ肷��Εs�v�ł͂��邪�A
'��ԍ��̃n�[�h�R�[�f�B���O�͂��Ȃ��ق�������

Private Type Head
    
    Caption As String
    Columns As Long

End Type

'���̃��W���[������J�����s�b�L���O�V�[�g�t�@�C�����N���[�Y����̂ɕK�v�ȃp�u���b�N�ϐ�
Public IsFileNewOpen As Boolean
Public PickingFileName As String

Function CheckPickingProducts(Optional IsMsgBox As Boolean = True)

'�e�X�g�p
'Sub CheckPickingProducts()

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
Dim LastRow As Integer
LastRow = WsPicking.Range("B1").SpecialCells(xlCellTypeLastCell).Row

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


'�s�b�L���O�V�[�g�̒����ԍ���A�����Җ���A���i�R�[�h��A���l�����肷��
'�w�b�_�[�z���p��
Dim Header(3) As Head
Header(0).Caption = "�����ԍ�"
Header(1).Caption = "�͂��於"
Header(2).Caption = "���i�R�[�h"
Header(3).Caption = "���P�[�V����" '���l�̎擾�Ɏg��

Dim h As Integer
For h = 0 To 3
    Header(h).Columns = WsPicking.Rows(1).Find(Header(h).Caption).Column
Next

'�����t�H�[����߂����߂ɋ󌟍�
WsPicking.Rows(1).Find ("")

'�s�b�L���O�V�[�g�̒����Җ��́u�����於�v�A���c�Ǘ��͒����Җ�
'���c�Ǘ��̖��O�ŒT���āA�Ȃ���΁u�s�b�L���O�V�[�g�Y�������v
Dim i As Integer
For i = 2 To LastRow
    
    Dim CurrentId As String
    CurrentId = WsPicking.Cells(i, Header(0).Columns).Value
    
    Dim CurrentBuyerName As String
    CurrentBuyerName = WsPicking.Cells(i, Header(1).Columns).Value
    
    Dim CurrentCode As String
    CurrentCode = WsPicking.Cells(i, Header(2).Columns).Value
        
    Dim CurrentNote As String
    CurrentNote = WsPicking.Cells(i, Header(3).Columns + 1).Value
        
    '�R�[�h�����t�[�̌`���ɕϊ� 012345->12345
    If CurrentCode Like "0#####" Then CurrentCode = Right(CurrentCode, 5)
    
    Dim o As order
    
    '�w�i�F�����ł͂Ȃ����i=�Z���^�[�݌ɗL��A�s�b�L���O�\
    If Not WsPicking.Cells(i, 1).Interior.Color = 16777215 Then
        
        '�����Җ�����A�ǂ̒����̏��i������
        Set o = FindByBuyerName(CurrentBuyerName, CurrentCode, TodaysOrders)
        
        '���c�Ǘ��V�[�g�͒����Җ��A�s�b�L���O�V�[�g�͈��掁���̂��߁A
        'FindByBuyerName���\�b�h�Œ����̈�v���Ƃꂸ��߂�l����̏ꍇ������
        
        If Not o Is Nothing Then

            o.Products(CurrentCode).IsPickingDone = True
                   
        End If
    
    End If
    
    'H��ɉ��������Ă遁����Ŕc�����Ă���݌ɏ�
    If CurrentNote <> "" Then
 
       '�����Җ�����A�ǂ̒����̏��i������
        Set o = FindByBuyerName(CurrentBuyerName, CurrentCode, TodaysOrders)
        
        '���̒����̏��i�I�u�W�F�N�g�Ƀs�b�L���O�̃t���O��o�^
        o.Products(CurrentCode).CenterStockState = CurrentNote

    End If
    
Next i

'TodaysOrder�̊e�����̊e���i��IsPickingDone��]�L

'Dim w As Variant
'Dim v As Variant

For Each v In TodaysOrders
    For Each w In TodaysOrders(v).Products

        If TodaysOrders(v).Products(w).IsPickingDone = False Then
            
            '�s�b�L���O�X�e�[�^�X���V�[�g�ɓ]�L�AFalse���Ɓu�Ȃ��v�{�{����z����
            'Todo:Order��Product�I�u�W�F�N�g��n���āAOrderSheet�ɔ��肵�ē]�L������
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

Private Function FindCol(Caption As String, PickingWorkSheet As Worksheet) As Integer



End Function
