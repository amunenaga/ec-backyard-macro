Attribute VB_Name = "ComputeQuantity"
'�I�����O�̎Г��V�X�e��-���݌ɂ̌����݌덷���A2��8���͂��̐��������Ă���݌ɐ����ۂ߂�
Public PreDecreaseQty As Long

Type Syokon
    Quantity As Long
    Status As String
    VenderCode As String
    
End Type
Sub BuildQtyCsv()

'�Г��V�X�e���̋敪�A���t�[�f�[�^��Abstract�A�݌Ɍ���V�[�g�A�p�ԃV�[�g���`�F�b�N���āA
'���t�[�ɃA�b�v���[����݌ɐ��AAllow-overdraft���Z�b�g���܂��B

'�Г��V�X�e���f�[�^��DL��OK���A�m�FPopup
go = MsgBox(prompt:="�Г��V�X�e���f�[�^���p�ӂł��Ă��܂����H" & vbLf & "�܂��A���t�[�f�[�^�̑O����s���̓N���A���Ă�낵���ł����H", Buttons:=vbYesNo)

If go <> vbYes Then
    MsgBox "�������I�����܂��B"
    End
End If

'�����O�̃f�[�^�`�F�b�N
If Not SyokonMaster.Cells(1, 1).Value = "Code" Then
    
    MsgBox "�Г��V�X�e���f�[�^�Ȃ��A�������I�����܂��B"
    Exit Sub

End If

If Not SecondInventry.Cells(1, 1).Value = "JAN" Then
    
    MsgBox "�l�b�g�p�݌Ƀf�[�^�Ȃ��A�������p�����܂����H"

End If

'���Ԍv�������܂�
'7��6000�s�̏�����420�b���炢

Dim startTime As Long
startTime = Timer

'����

Call FetchSecondInventry

Call FetchYahooCSV

'�e�V�[�g�̃R�[�h�͈͂𖼑O�ŌĂяo����悤�Ē�`
Call SetRangeName


'---��������---

'�Г��V�X�e���f�[�^����S�A�C�e���ɍ݌ɂ��Z�b�g
Call SetQuantity

'���t�[�f�[�^�V�[�g����CSV��ۑ�
Call PutQtyCsv

'�I������������
Dim endTime As Long
endTime = Timer

Call ApendProcessingTime(endTime - startTime)

MsgBox "���s���ԁF" & endTime - startTime & " �b"

End Sub

Sub buildSecondInventryQty()

Dim startTime As Long
startTime = Timer

'����

Call FetchSecondInventry
'�e�V�[�g�̃R�[�h�͈͂𖼑O�ŌĂяo����悤�Ē�`
Call SetRangeName

'�Г��V�X�e���f�[�^����S�A�C�e���ɍ݌ɂ��Z�b�g
Call SetQuantity

'�I������������
Dim endTime As Long
endTime = Timer

Call ApendProcessingTime(endTime - startTime)
With yahoo6digit

    .Activate
       
    Dim StatusCol As Integer
    StatusCol = .Rows(1).Find("status").Column
    
    .Range("A1").CurrentRegion.AutoFilter Field:=StatusCol, Criteria1:=Array( _
            "�l�b�g�p�݌ɂɗL", "�l�b�g�p�݌Ɋ���" _
            ), Operator:=xlFilterValues
    
    '�t�B���^�[���������W���Z�b�g
    Dim A As Range
    Set A = .Range("C1").CurrentRegion.SpecialCells(xlCellTypeVisible)
    
    Dim B As Range
    Set B = .Range("C2").Resize(Range("C1").SpecialCells(xlCellTypeLastCell).row - 1, 1)
    
    Dim CodeRange As Range
    Set CodeRange = Application.Intersect(A, B)

End With

'�����o���pCSV�V�[�g��p��
Worksheets("CSV").Cells.Clear

'�w�b�_�[����������
header = Array("code", "quantity", "allow-overdraft")

Worksheets("CSV").Range("A1:C1") = header

Worksheets("���t�[�f�[�^").Activate

colQuantity = yahoo6digit.Rows(1).Find("quantity").Column
colAllow = yahoo6digit.Rows(1).Find("allow-overdraft").Column

Dim i As Long
i = 2
'�R�[�h�����W�ɑ΂��āAr.row�ōs�ԍ������o���ē����s��Quantity/Allow�̒l���擾����
For Each r In CodeRange

    Code = r.Value

    Qty = Cells(r.row, colQuantity).Value
    pur = Cells(r.row, colAllow).Value

    Worksheets("CSV").Range("A" & i & ":C" & i) = Array(Code, Qty, pur)

    i = i + 1

Next

Worksheets("CSV").Activate

'CSV�ǋL����
Dim FSO As New FileSystemObject
Dim Csv As Object

'�ǋL���[�h ForAppending �Ńt�@�C�����J��
Set Csv = FSO.OpenTextFile(Filename:=ThisWorkbook.Path & "\" & "���t�[�݌ɍX�V0413.csv", IOMode:=8)


For i = 94 To 709
    
    With Worksheets("Csv")
        Csv.WriteLine (CStr(.Cells(i, 1).Value) & "," & CStr(.Cells(i, 2).Value) & "," & CStr(.Cells(i, 3).Value))
    End With

Next

MsgBox "���s���ԁF" & endTime - startTime & " �b"

End Sub

Sub SetQuantity()

'PreDecreaseQty�F�I�����O�̎Г��V�X�e��-���݌ɂ̌����݌덷��
'2��8���͎Г��V�X�e������5�������Ă���0.6�|�������݌ɐ����Z�b�g����

PreDecreaseQty = 0

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
        i = r.row
        
        
        'Debug.Assert i < 2902

        
        '�݌ɐݒ菜�O�V�[�g�ɂ���΁A�ȉ��̏����͍s��Ȃ��AContinue�֔��
        If WorksheetFunction.CountIf(ExceptQty.Range("ExceptCodeRange"), item.Code) > 0 Then GoTo continue
        
        'Abstract���E���B16�N1������ �S���i�Ŏg���Ă���̂ŋL�ڗL������͂Ȃ�
        item.Abstract = yahoo6digit.Cells(i, colAbstract).Value

               
        '�Г��V�X�e���V�[�g����Г��V�X�e���̒l���擾�A�o�^������Sy.Quantity=0
        Dim sy As Syokon
        sy = SyokonMaster.GetSyokonQtyKubun(item.Code)
        
        'Item�I�u�W�F�N�g�ɎГ��V�X�e���̒l���Z�b�g
        item.Status = sy.Status
        item.VenderCode = sy.VenderCode
        
        '�l�b�g�p�݌ɁA�p�ԁA�݌Ɍ���̊e�V�[�g���`�F�b�N
        item.CheckSecondInventry
        item.CheckEol
        item.CheckStockOnly
        
        
        '�ݒ�݌ɐ����Z�o�Z�b�g�A�l�b�g�p�݌�/�Г��V�X�e�����X�C�b�`���ēn���B
        
        If item.IsSecondInventry Then
            item.SetQuantity (item.SecondInventryQuantity)
        Else
            item.SetQuantity (sy.Quantity)
        End If
        
        '�����ۂ��Z�b�g
        item.SetAvailablePurchase
        
        '�Z�o�����݌ɂƁA���肵��Allow-overDraft�������o��
        
        .Cells(i, colQuantity).Value = item.Quantity
        
        If item.AvailablePurchase Then  'Allow-overdraft��Bool�l�Ȃ̂�1/0�ɒu�������ďo��
            .Cells(i, colAllow).Value = 1
        Else
            .Cells(i, colAllow).Value = 0
        End If
        
        .Cells(i, colStatus).Value = item.Status
       
continue:
    
       Set item = Nothing
       
    Next r

End With


'�ꎞ��~���㏑��
Call halt.setHalt

'�ݔp�A������0�͔p�ԁE�I���ֈړ�
Call CheckEolInStockOnly

'����̃l�b�g�p�݌ɃA�C�e���̍݌ɗL���ʃV�[�g�ɃR�s�[���Ă���
Call StackLastQty

'������s���ɎГ��V�X�e���f�[�^DL���ĂȂ��ƃG���[�o��悤�A�Г��V�X�e���̍���f�[�^���ړ�

SyokonMaster.Range("A1").CurrentRegion.Cut Destination:=SyokonMaster.Range("K1")


End Sub

Sub UpdateSecondInventryQty()

Call FetchSecondInventry

End Sub

Sub PutQtyCsv()
'FileSystemObject�̃e�L�X�g�X�g���[����CSV�t�@�C���𐶐����āATextStream�œ��e�𗬂����݂܂��B
'���b�ŏI���܂��B

With yahoo6digit '���t�[�f�[�^�̉�����

    .Activate
    
    '�u"�o�^�Ȃ�"�v�Ɓu"��"�v ����2�ȊO���t�B���^�[�ŕ\���cTODO�F1��ڂ���t�B���^�[�̏󋵂��`�F�b�N��������������

    
    Dim StatusCol As Integer
    StatusCol = .Rows(1).Find("status").Column
    
    .Range("A1").CurrentRegion.AutoFilter Field:=StatusCol, Criteria1:=Array( _
             "�݌ɔp��", "�݌ɏ���", "�l�b�g�p�݌ɗL", "�l�b�g�p�݌Ɋ���", "��������", "�o�^�̂�", "�p�ԕi", "�̔����~", "�W��" _
            ), Operator:=xlFilterValues
    
    '�t�B���^�[���������W���Z�b�g�ACSV�̃w�b�_�[�͕ʓr��������ł����̂ŁA2�s�ڈȍ~�̃����W�B
    Dim A As Range
    Set A = .Range("C1").CurrentRegion.SpecialCells(xlCellTypeVisible)
    
    Dim B As Range
    Set B = .Range("C2").Resize(Range("C1").SpecialCells(xlCellTypeLastCell).row - 1, 1)
    
    Dim CodeRange As Range
    Set CodeRange = Application.Intersect(A, B)

End With

'�����o���pCSV��p��
Dim day As String
day = Format(Date, "mm") & Format(Date, "dd")

Dim OutputCsvName As String
OutputCsvName = "�Г��V�X�e���݌ɃA�b�v�p" & day & ".csv"

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
    
    Qty = Cells(r.row, colQuantity).Value
    pur = Cells(r.row, colAllow).Value
    
    TS.WriteLine Code & "," & Qty & "," & pur

Next

TS.Close

End Sub


Sub StackLastQty()
'�O��l�b�g�p�݌ɂ��N���A�[
'���t�[�f�[�^����l�b�g�p�݌ɂɗL���t�B���^�[���đO��l�b�g�p�݌ɂɃR�s�[
LastSecondInventry.Cells.Clear

yahoo6digit.Activate

Dim StatusCol As Integer
StatusCol = Rows(1).Find("status").Column

Range("A1").CurrentRegion.AutoFilter Field:=StatusCol, Criteria1:="�l�b�g�p�݌ɂɗL"

'�t�B���^�[�����\���̈�
Dim A As Range
Set A = Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible)

'���t�[�f�[�^�S�̂̃����W
Dim B As Range
Set B = yahoo6digit.AutoFilter.Range

'AB�̌��������W���f�[�^�����W�Ƃ��ăZ�b�g�A�l�b�g�p�݌ɂɗL���i�̃����W
Dim InSecondInventryRange As Range
Set InSecondInventryRange = Application.Intersect(A, B)

InSecondInventryRange.Copy LastSecondInventry.Range("A1")

End Sub
