Attribute VB_Name = "PutYahooInvoiceFile"
Option Explicit

Sub �����ԍ��ꊇ�t�@�C���쐬()
'���c�Ǘ��V�[�g�̖����������ɁA�o�׏ڍ�CSV���珈�����шꊇ�A�b�v���[�h�pCSV�𐶐����܂��B
'2015�N5��19���쐬

'2015/6/2�A�N���X���W���[�����g���ă��t�@�N�^�����O���܂���
'2015/6/3�A�����ԍ����o�׏ڍ׈ꗗ.xlsx����擾����悤�ɕύX���܂����B
'2015/10/22 �i���\����t���܂����B�^�`�F�b�N�����܂��B
'2015/12/02 �o�גʒm���O�@�\�͂Ȃ��Ă�������������Ȃ��B�����݌ɂ𔻒肵�āA����̑����V�X�e���ւ̑��荞�݃f�[�^�쐬���Ă�炵���̂ŁB
            '����Ԉ���đ����쐬�����Ƃ�����ǂ��ł��`�F�b�N����Ă��Ȃ��̂Ť�o�גʒm���O�͈��������@�\�����܂��
'2016/1/6 �����̏����݌ɂ��Ȃ���Α����N�[���Ȃ��}�N���@�\���ĂȂ���������B
            
'-------------------------------------�؂����-----------------------------------------------

OpPanel.Hide

Application.ScreenUpdating = False

'�v���O���X�o�[�̏��� �\�����@�͂��̃T�C�g���悭�܂Ƃ܂��Ă� http://hideprogram.web.fc2.com/vba/vba_ProgressBarForm.html
ShippingFileProgress.ProgressBar.Min = 1
ShippingFileProgress.ProgressBar.Max = 5

'�i���E�B���h�E�̏󋵕\�����Z�b�g
ShippingFileProgress.ProgressBar.Value = 1
ShippingFileProgress.ShowCurrentProcess.Caption = "�󒍎捞/�s�b�L���O�捞 �`�F�b�N��"

'�i���E�B���h�E��\�� ���[�h���X�w�肾�ƃo�b�N�O���E���h�ŏ������i��
ShippingFileProgress.Show vbModeless

'�{���̎󒍁A�s�b�L���O�V�[�g���]�L�ς��`�F�b�N���܂��B

If LogSheet.Range("LastUpdatePickingSheet").Value <> Date Then

    On Error Resume Next '�s�b�L���O�V�[�g�t�@�C�����J���Ȃ��Ă����s�A�����t�@�C�������ɕK�{�ł͂Ȃ��̂�
        
        Call CheckPickingProducts(IsMsgBox:=False)
   
    On Error GoTo 0
    
    LogSheet.Range("B9").Value = Date
    
    If FetchPickingSheet.IsFileNewOpen Then Workbooks(FetchPickingSheet.PickingFileName).Close

End If

'�{�����̑����ԍ��z����쐬 ��J������ShippingFileProgress�̍X�V������Invoces���ł���Ă܂�
Dim TodaysInvoices As Invoices
Set TodaysInvoices = New Invoices

TodaysInvoices.fetchReportXlsx

ShippingFileProgress.ProgressBar.Value = 3
ShippingFileProgress.ShowCurrentProcess.Caption = "���������� ���X�g�쐬��"

'�����������̔z����쐬
Dim TodaysUndispatch As Dictionary
Set TodaysUndispatch = OrderSheet.getUndispatchOrders

'����������dictionary���o���Ă��邩�`�F�b�N
If TodaysUndispatch.Count = 0 Then
    
    MsgBox prompt:="���o�ג�����0���ł��B" & vbLf & "���c�Ǘ��V�[�g���m�F���Ă��������B" & vbLf _
                    & "Dictionary.count = 0 in ""OrderSheet""" _
                    , Buttons:=vbExclamation
    End

End If

ShippingFileProgress.ProgressBar.Value = 4
ShippingFileProgress.ShowCurrentProcess.Caption = "�����������̑����ԍ���]�L��"

Dim TodaysShipping As ShippingOrders
Set TodaysShipping = New ShippingOrders

Call TodaysShipping.createShippingList(TodaysUndispatch, TodaysInvoices)

ShippingFileProgress.ProgressBar.Value = 5
ShippingFileProgress.ShowCurrentProcess.Caption = "�ꊇ�A�b�v���[�h�p�t�@�C����ۑ���"

TodaysShipping.putCsv

ShippingFileProgress.Hide

ThisWorkbook.Save

Call ������̋󗓂̂ݕ\��

Application.ScreenUpdating = True

MsgBox prompt:="���t�[�����ԍ��ꊇ" & Format(Date, "mmdd") & "   �ۑ����܂����B" & vbLf _
                & "�䂤�p�P�b�g�������͎蓮�œ��͂����肢���܂��B" _
                , Buttons:=vbInformation

End

End Sub
