Attribute VB_Name = "ComputeQuantity"
Type Syokon
    Quantity As Long
    Status As String
    VenderCode As String
    
End Type
Sub PutQtyCsv()

'�����̋敪�A���t�[�f�[�^��Abstract�A�݌Ɍ���V�[�g�A�p�ԃV�[�g���`�F�b�N���āA
'���t�[�ɃA�b�v���[����݌ɐ��AAllow-overdraft���Z�b�g���܂��B

If Not SecondInventry.Cells(1, 1).Value = "JAN" Then
    
    MsgBox "�I���f�[�^�Ȃ��A�������p�����܂����H"

End If

'���Ԍv�������܂�

Dim startTime As Long
startTime = Timer

'����

Call FetchSecondInventry

'Call FetchYahooCSV

'�e�V�[�g�̃R�[�h�͈͂𖼑O�ŌĂяo����悤�Ē�`
Call SetRangeName


'---��������---

'�����f�[�^����S�A�C�e���ɍ݌ɂ��Z�b�g
Call SetQuantity

'���t�[�f�[�^�V�[�g����CSV��ۑ�
'Call PutQtyCsv

'�I������������
Dim endTime As Long
endTime = Timer

Call ApendProcessingTime(endTime - startTime)

MsgBox "���s���ԁF" & endTime - startTime & " �b"

End Sub

Sub buildSecondInventry()

Dim startTime As Long
startTime = Timer

'����

Call FetchSecondInventry
'�e�V�[�g�̃R�[�h�͈͂𖼑O�ŌĂяo����悤�Ē�`
Call SetRangeName

'�����f�[�^����S�A�C�e���ɍ݌ɂ��Z�b�g
Call SetQuantity

'�I������������
Dim endTime As Long
endTime = Timer

Call ApendProcessingTime(endTime - startTime)

MsgBox "���s���ԁF" & endTime - startTime & " �b"

End Sub

Sub SetQuantity()

'PreDecreaseQty�F�I�����O�̏���-���݌ɂ̌����݌덷��
'2��8���͏�������5�������Ă���0.6�|�������݌ɐ����Z�b�g������

yahoo6digit.Activate

Application.ScreenUpdating = False


'�w�b�_�[��ǋL
yahoo6digit.Range("A1").End(xlToRight).Offset(0, 1).Resize(1, 3) = Array("quantity", "allow-overdraft", "status")

Dim item As New item

Dim colAbstract As Integer, colQuantity As Integer, colAllow As Integer, colStatus As Integer '�Ȃ񂼂����ƃX�}�[�g�Ȓ�`���@�����邾��
colAbstract = yahoo6digit.Rows(1).Find("abstract").Column

colQuantity = yahoo6digit.Rows(1).Find("quantity").Column
colAllow = yahoo6digit.Rows(1).Find("allow-overdraft").Column
colStatus = yahoo6digit.Rows(1).Find("status").Column

Dim r As Range

With yahoo6digit 'With�\�����ł̓I�u�W�F�N�g�Q�Ƃ��J��Ԃ���Ȃ����߁A�������������ɂȂ�炵��

    For Each r In .Range("YahooCodeRange")
        
        Set item = New item
        item.Code = r.Value
        
        Dim i As Long  'TODO:�s�ԍ����i�[����i�͗v��Ȃ��̂ł́c
        i = r.Row
        
        'Debug.Assert i < 1000
        
        '�݌ɐݒ菜�O�V�[�g�ɂ���΁A�ȉ��̏����͍s��Ȃ��AContinue�֔��
        If WorksheetFunction.CountIf(ExceptQty.Range("ExceptCodeRange"), item.Code) > 0 Then GoTo Continue
        
        'Abstract���E���B16�N1������ �S���i�Ŏg���Ă���̂ŋL�ڗL������͂Ȃ�
        item.Abstract = yahoo6digit.Cells(i, colAbstract).Value

               
        '�����V�[�g���珤���̒l���擾�A�o�^������Sy.Quantity=0
        Dim sy As Syokon
        sy = SyokonMaster.GetSyokonQtyKubun(item.Code)
        
        'Item�I�u�W�F�N�g�ɏ����̒l���Z�b�g
        item.Status = sy.Status
        item.VenderCode = sy.VenderCode
        
        '�I�Ȃ��A�p�ԁA�݌Ɍ���̊e�V�[�g���`�F�b�N
        item.CheckSecondInventry
        item.CheckEol
        item.CheckStockOnly
        
        '�ݒ�݌ɐ��̃Z�b�g�A�݌ɐ��̓X�����X�Ɉ�{��
        
        If Slims.HasLocation(item.Code) Then
            
            item.Quantity = Slims.getQuantity(item.Code)
        
        Else
            
            item.Quantity = 0
        
        End If
        
        '��z�ۂ��Z�b�g
        item.SetAvailablePurchase
        
        '�Z�o�����݌ɂƁA���肵��Allow-overDraft�������o��
        
        'Debug.Assert item.Quantity > 0
        
        .Cells(i, colQuantity).Value = item.Quantity
        
        If item.AvailablePurchase Then  'Allow-overdraft��Bool�l�Ȃ̂�1/0�ɒu�������ďo��
            .Cells(i, colAllow).Value = 1
        Else
            .Cells(i, colAllow).Value = 0
        End If
        
        .Cells(i, colStatus).Value = item.Status
       
Continue:
    
       Set item = Nothing
       
    Next r

End With


'�ꎞ��~���㏑��
Call halt.setHalt

'�ݔp�A������0�͔p�ԁE�I���ֈړ�
Call CheckEolInStockOnly

'����̒I�����A�C�e���̍݌ɗL���ʃV�[�g�ɃR�s�[���Ă���
Call StackLastQty

End Sub

Sub UpdateSecondInventryQty()

Call FetchSecondInventry

End Sub

Sub PutCsv()
'FileSystemObject�̃e�L�X�g�X�g���[����CSV�t�@�C���𐶐����āATextStream�œ��e�𗬂����݂܂��B
'���b�ŏI���܂��B

With yahoo6digit '���t�[�f�[�^�̉�����

    .Activate
    
    '�u"�o�^�Ȃ�"�v�Ɓu"��"�v ����2�ȊO���t�B���^�[�ŕ\���cTODO�F1��ڂ���t�B���^�[�̏󋵂��`�F�b�N��������������
    '16-2-29 �p�Ԃ̋敪���u���p�ԁv�ɂȂ�܂����B
    
    Dim StatusCol As Integer
    StatusCol = .Rows(1).Find("status").Column
    
    .Range("A1").CurrentRegion.AutoFilter Field:=StatusCol, Criteria1:=Array( _
            "�r�o����", "����i", "�݌ɔp��", "�݌ɏ���", "�I�Ȃ��ɗL", "�I�Ȃ�����", "��������", "�o�^�̂�", "���p�ԕi", "�̘H����", "�̔����~", "�W��" _
            ), Operator:=xlFilterValues
    
    '�t�B���^�[���������W���Z�b�g�ACSV�̃w�b�_�[�͕ʓr��������ł����̂ŁA2�s�ڈȍ~�̃����W�B
    Dim A As Range
    Set A = .Range("C1").CurrentRegion.SpecialCells(xlCellTypeVisible)
    
    Dim B As Range
    Set B = .Range("C2").Resize(Range("C1").SpecialCells(xlCellTypeLastCell).Row - 1, 1)
    
    Dim CodeRange As Range
    Set CodeRange = Application.Intersect(A, B)

End With

'�����o���pCSV��p��
Dim day As String
day = Format(Date, "mm") & Format(Date, "dd")

Dim OutputCsvName As String
OutputCsvName = "�����݌ɃA�b�v�p" & day & ".csv"

Dim FSO As Object 'TODO:���O�o�C���f�B���O�ɕύX
Set FSO = CreateObject("Scripting.FileSystemObject")

Dim TS As Object
    
Set TS = FSO.CreateTextFile(Filename:=ThisWorkbook.Path & "\" & OutputCsvName, _
                            OverWrite:=True)
                            
'�w�b�_�[����������
header = "code,quantity,allow-overdraft"

TS.WriteLine header

colQuantity = yahoo6digit.Rows(1).Find("quantity").Column
colAllow = yahoo6digit.Rows(1).Find("allow-overdraft").Column
colStatus = yahoo6digit.Rows(1).Find("status").Column

'�R�[�h�����W�ɑ΂��āAr.row�ōs�ԍ������o���ē����s��Quantity/Allow�̒l���擾����
For Each r In CodeRange
    
    Code = r.Value
    
    qty = Cells(r.Row, colQuantity).Value
    pur = Cells(r.Row, colAllow).Value
    
    TS.WriteLine Code & "," & qty & "," & pur

Next

TS.Close

End Sub


Sub StackLastQty()
'�O��I�������N���A�[
'���t�[�f�[�^����I�Ȃ��ɗL���t�B���^�[���đO��I�����ɃR�s�[
LastSecondInventry.Cells.Clear

yahoo6digit.Activate

Dim StatusCol As Integer
StatusCol = Rows(1).Find("status").Column

Range("A1").CurrentRegion.AutoFilter Field:=9, Criteria1:="�I�Ȃ��ɗL"

'�t�B���^�[�����\���̈�
Dim A As Range
Set A = Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible)

'���t�[�f�[�^�S�̂̃����W
Dim B As Range
Set B = yahoo6digit.AutoFilter.Range

'AB�̌��������W���f�[�^�����W�Ƃ��ăZ�b�g�A�I�Ȃ��ɗL���i�̃����W
Dim InSecondInventryRange As Range
Set InSecondInventryRange = Application.Intersect(A, B)

InSecondInventryRange.Copy LastSecondInventry.Range("A1")

End Sub
