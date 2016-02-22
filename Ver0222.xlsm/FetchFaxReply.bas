Attribute VB_Name = "FetchFaxReply"
Option Explicit

'FAX�[���񓚃��X�g�̂���f�B���N�g���A�Ō�K��\�}�[�N
Const PURCHASE_LOG_FOLDER As String = "\\Server02\���i��\�l�b�g�̔��֘A\�����֘A\�����������o�b�N�A�b�v\"
Const REPLY_XLSM_FILE As String = "FAX�[���񓚃��X�g.xlsm"
Const REPLY_SHEET_NAME As String = "�[�����X�g"

Dim RangeFaxReplyCode As Range

'Sub FetchFaxReply() '�P�̃e�X�g�̂��߂�Sub�؂�ւ��R�����g�A�E�g�AFunction�ɂ��Ă�ƃ}�N���ɕ\�����Ȃ��̂�
Function FetchFaxReply(Optional arg As Variant = "")

'���c�u�A�C�e���v���X�g�����A�����ԍ��E�R�[�h�E������
'�ԐMFAX�̃t�@�C�����J��
'�R�[�h��Find�����������ĒT��
'Y��+������΃��t�[�̒�����
'���t�łȂ������[�J�[�󋵂ɓ]�L�A���ד�������Γ]�L

Application.ScreenUpdating = False

'��z�����iDictionary���쐬
Dim CurrentPurchase As Dictionary
Set CurrentPurchase = OrderSheet.getCurrentPurchase

'FAX�[���񓚃��X�g���J��


Dim wb As Workbook
Dim WsFaxReply As Worksheet

'���[�N�u�b�N���J���Ă���΂�����g��
For Each wb In Workbooks
    
    If wb.Name = REPLY_XLSM_FILE Then
    
        Set WsFaxReply = wb.Sheets(REPLY_SHEET_NAME)

    End If
    
Next wb

'���[�N�u�b�N���J���ăZ�b�g
If WsFaxReply Is Nothing Then

    Set wb = Workbooks.Open(PURCHASE_LOG_FOLDER & REPLY_XLSM_FILE)
    Set WsFaxReply = wb.Sheets(REPLY_SHEET_NAME)
            
End If

Workbooks(REPLY_XLSM_FILE).Activate

'If WsFaxReply.AutoFilterMode = True Then WsFaxReply.Range("A1").AutoFilter

'FAX�ԐM���X�g�̏��i�R�[�h�����W

Set RangeFaxReplyCode = WsFaxReply.Range("I2").Resize(WsFaxReply.Range("I2").CurrentRegion.Rows.Count, 1)

'FAX�[���񓚃��X�g���J���āA��z�ςݏ��i���X�g���擾����
Dim v As Variant

For Each v In CurrentPurchase
    
    Dim FirstFoundCell As Range
    Set FirstFoundCell = RangeFaxReplyCode.Find(CurrentPurchase(v).Code)
       
    '���c���X�g�ɊY�����i�R�[�h���Ȃ���Ύ���Product��
    If FirstFoundCell Is Nothing Then GoTo continue
                  
    Call FindArrivalDate(CurrentPurchase(v), FirstFoundCell)
        
continue:

Next v

'For Each v In CurrentPurchase
'    Debug.Print CurrentPurchase(v).OrderId & ":" & CurrentPurchase(v).Code
'Next v

For Each v In CurrentPurchase
    Call OrderSheet.WriteEstimatedArrivalDate(CurrentPurchase(v))
Next v

'FAX�[���񓚃��X�g�����A�J�����ςȂ����ƃG�N�Z�����d������B
'2015-09-15���_�Ńt�@�C����4MB���炢����
Workbooks(REPLY_XLSM_FILE).Close SaveChanges:=False

Call �������̂ݕ\��

ThisWorkbook.Save

Application.ScreenUpdating = True

MsgBox prompt:="�ԐM���X�g�Ǎ�����"

End

End Function

Private Sub FindArrivalDate(Product As Product, FoundCell As Range)
'�P��Finder
'�ԐMFAX�̕ԐM�L�ڗ���n�[�h�R�[�f�B���O���Ă���̂ŁA����
'���� FoundCell�̃����W�ł����̂��낤���H�Ȃ񂩕�

Dim FirstFoundCellAddress As String
FirstFoundCellAddress = FoundCell.Address

Do
    
    'Debug.Assert FoundCell.Address <> "$I$252" '����A�h���X�̎��ɒ�~
    'Debug.Print FoundCell.Address
    
    Dim PurDate As String, Identifier As String, VenderReply As Variant, EstimatedArrivalDate As Variant
    
    PurDate = CStr(Range("F" & FoundCell.Row).Value)
    Identifier = Range("E" & FoundCell.Row).Value
    VenderReply = Range("W" & FoundCell.Row).Value
    EstimatedArrivalDate = Range("Y" & FoundCell.Row).Value
            
    '���t�[���c��FAX�ԐM���X�g�̏��i����v����Ƃ݂Ȃ�����
    If PurDate = Format(Product.PurchaseDate, "mdd") Then '����������v����
        
        If InStr(Identifier, "Y") > 0 Or InStr(Identifier, "+") > 0 Then '���[�����ʎq���uY���܂ށv���u+���܂ށv
            
            '���[�J�[�ԐM���e�����t�łȂ��ꍇ�A���[�J�[�󋵃v���p�e�B�ɓ]�L
            If Not IsDate(VenderReply) Then 'IsDate�֐��͓��t�^�ɕϊ��\���𔻒肷��炵���A�����Ȍ^�����ł͂Ȃ�
                Product.VenderStatus = VenderReply
            End If
            
            '���ח\��������t�^�Ȃ�A���ח\��v���p�e�B�ɓ]�L
            If IsDate(EstimatedArrivalDate) Then
                Product.EstimatedArrivalDate = EstimatedArrivalDate
            End If
            
            Exit Do
            
        End If
    
    End If
    
    '������������A�ŏ��̌����s�ƈ�v����܂�Loop�p��
    Set FoundCell = RangeFaxReplyCode.FindNext(FoundCell)
    
    If FoundCell Is Nothing Then Exit Do
    
Loop Until FirstFoundCellAddress = FoundCell.Address

'�Q�Ɠn���ŃI�u�W�F�N�g������Ă�̂ŁA�l�̕ԋp�͕s�v

End Sub
