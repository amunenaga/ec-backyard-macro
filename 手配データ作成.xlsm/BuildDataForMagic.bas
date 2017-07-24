Attribute VB_Name = "BuildDataForMagic"
Option Explicit
Const OPERATOR_CODE As Integer = 329

Type Purchase
'��z���ʓ��̓V�[�g1�s���ɑ������郆�[�U�[��`�^

    Code As String
    ProductName As String
    
    VendorCode As Long
    VendorName As String
    
    UnitCost As Long
    
    PurchaseQuantity As Long
    RequireQuantity As Long
    
    RequireMallCount As String
    
    WarehouseNumber As Integer
    
    IsPickup As Integer
    
    IsHold As Boolean
    HoldReason As String

End Type

Sub BuildPurcahseData()
'�������ʌ���V�[�g�����ɁA�t�@�C���������o��
'�u�����V�X�e���p�f�[�^�o�́v�{�^���ŌĂяo�����

'�e�V�[�g���󂩃`�F�b�N
Dim Sh As Variant
For Each Sh In Array(Worksheets("Magic�ꊇ�o�^"), Worksheets("Magic����͗p"), Worksheets("�������i���X�g"), Worksheets("�ۗ�"))
    Call PrepareSheet(Sh)
Next

'�f�[�^�o�͗p�̃V�[�g�ɁA1�s���R�s�[
Worksheets("��z���ʓ��̓V�[�g").Activate

Dim i As Long
For i = 2 To Range("A1").End(xlDown).Row

    Dim CurrentPurchase As Purchase
    CurrentPurchase = ReadPurchase(i)
    
    If CurrentPurchase.IsHold Then
        Call WriteHoldList(CurrentPurchase)
    Else
        
        Call WriteBackupSheet(CurrentPurchase)
        
        If CurrentPurchase.Code Like "######" Then
            Call WriteMagicTxt(CurrentPurchase)
        Else
            Call WriteMagicManualInput(CurrentPurchase)
        End If
        
    End If

Next

Worksheets("Magic�ꊇ�o�^").Columns("A:E").AutoFit
Worksheets("Magic����͗p").Columns("A:I").AutoFit

'�o�͗p�V�[�g���t�@�C���Ƃ��ĕۑ����Ă���

'Magic�ꊇ�o�^�V�[�g��V�K�u�b�N�ɃR�s�[�A�g���q.txt�A�J���}��؂�A�w�b�_�[�����ŕۑ�
Worksheets("Magic�ꊇ�o�^").Copy
ActiveSheet.Rows(1).Delete

Dim FileName As String
FileName = "\Magic�o�^�p" & Format(Date, "MMdd") & ".txt"

If Dir(ThisWorkbook.path & FileName) <> "" Then
    FileName = Replace(FileName, Format(Date, "MMdd"), Format(Date, "MMdd") & "-" & Format(Time, "hhmm"))
End If

Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs FileName:=ThisWorkbook.path & FileName, FileFormat:=xlCSV
    ActiveWorkbook.Close

'�o�b�N�A�b�v��ۑ�
ThisWorkbook.Worksheets("�������i���X�g").Copy

With ActiveSheet
    .Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
    .Rows(1).Insert
    .Range("B1").Value = "�ޯ����ߓ��� : " & Format(Date, "YYYY/MM/dd") & " " & Format(Time, "hh:mm:ss")
End With

ActiveWorkbook.SaveAs FileName:="\\Server02\���i��\�l�b�g�̔��֘A\�����֘A\�����������o�b�N�A�b�v\BU" & Format(Date, "YYYYMMDD") & "-" & Format(Now, "hhmmss") & ".xlsx"
ActiveWorkbook.Close

'�ۗ���ۑ�
Worksheets("�ۗ�").Copy

FileName = "\�ۗ�" & Format(Date, "MMdd") & ".xlsx"

If Dir(ThisWorkbook.path & FileName) <> "" Then
    FileName = Replace(FileName, Format(Date, "MMdd"), Format(Date, "MMdd") & "-" & Format(Time, "hhmm"))
End If

ActiveWorkbook.SaveAs FileName:=ThisWorkbook.path & FileName

'c�ۗ��֒ǋL���Ă������
Call AppendHoldPurWokbook(ActiveWorkbook)

ActiveWorkbook.Close

'Magic���͗pExcel�t�@�C����ۑ�
Sheets(Array("Magic�ꊇ�o�^", "Magic����͗p")).Copy

FileName = "\Magic���̓f�[�^" & Format(Date, "MMdd") & ".xlsx"

If Dir(ThisWorkbook.path & FileName) <> "" Then
    FileName = Replace(FileName, Format(Date, "MMdd"), Format(Date, "MMdd") & "-" & Format(Time, "hhmm"))
End If

ActiveWorkbook.SaveAs FileName:=ThisWorkbook.path & FileName
ActiveWorkbook.Close

'�t�@�C���o�͊����A���̃u�b�N��ۑ�
ThisWorkbook.Save

Application.DisplayAlerts = True

MsgBox Prompt:="�t�@�C���ۑ����������܂����B", Buttons:=vbInformation

End Sub

Private Function ReadPurchase(ByVal Row As Long) As Purchase
'��z���ʓ��̓V�[�g����1�s��1�ϐ��ɓǂݍ���

Dim TmpPur As Purchase

With TmpPur
    .Code = Cells(Row, 7).Value  '�������̏��i�R�[�h�AJAN��6�P�^
    .ProductName = Cells(Row, 8).Value '���i���AJAN��z���̂ݕK�{
    
    .VendorCode = Cells(Row, 4).Value '��z��R�[�h
    .VendorName = Cells(Row, 5).Value '��z�於��
     
    .WarehouseNumber = IIf(Cells(Row, 6).Value = "V", "187", "180")  '�q�ɔԍ�

    .RequireQuantity = Cells(Row, 9).Value '��z�˗�����
    
    .RequireMallCount = Cells(Row, 6).Value '���[���ʂ̈˗�����

    '�����ۗ��ɊY�����邩�`�F�b�N���āA�t���O�𗧂Ă�

    '�ۗ��Ƃ��Đ��ʂ𒍈ӏ����ŏ㏑�����Ă��邩�H
    If IsNumeric(Cells(Row, 1).Value) Then
        TmpPur.PurchaseQuantity = Cells(Row, 1).Value
    Else
        TmpPur.HoldReason = Cells(Row, 1).Value
        TmpPur.IsHold = True
    End If
    
    .UnitCost = Cells(Row, 10).Value
    
    '����Ŏ�z���邩
    If Trim(Cells(Row, 11).Value) = "" Then
        .IsPickup = 2
    Else
        .IsPickup = Cells(Row, 11).Value
    End If
        
End With

ReadPurchase = TmpPur

End Function

Private Sub WriteMagicTxt(ByRef Purchase As Purchase)
    
    Dim WriteRow As Long, TargetSheet As Worksheet, Record As Variant
    
    With Purchase
        Record = Array( _
                    .WarehouseNumber, _
                    .Code, _
                    .PurchaseQuantity, _
                    .IsPickup, _
                    OPERATOR_CODE _
                    )
    End With
    
    Set TargetSheet = Worksheets("Magic�ꊇ�o�^")
    WriteRow = TargetSheet.Range("A1").SpecialCells(xlLastCell).Row + 1
    
    With TargetSheet
        .Cells(WriteRow, 2).NumberFormatLocal = String(9, "0")
        .Cells(WriteRow, 3).NumberFormatLocal = String(8, "0")
    
        .Cells(WriteRow, 1).Resize(1, 5).Value = Record
    End With
    
End Sub

Private Sub WriteMagicManualInput(ByRef Purchase As Purchase)

    Dim WriteRow As Long, TargetSheet As Worksheet, Record As Variant
    
    With Purchase
        Record = Array( _
                    .WarehouseNumber, _
                    .VendorCode, _
                    .VendorName, _
                    .Code, _
                    .ProductName, _
                    .PurchaseQuantity, _
                    .UnitCost, _
                    .IsPickup, _
                    OPERATOR_CODE _
                    )
    End With
    
    Set TargetSheet = Worksheets("Magic����͗p")
    WriteRow = TargetSheet.Range("A1").SpecialCells(xlLastCell).Row + 1
    
    TargetSheet.Cells(WriteRow, 4).NumberFormatLocal = "@"
    TargetSheet.Cells(WriteRow, 1).Resize(1, 9).Value = Record
    
End Sub

Private Sub WriteHoldList(ByRef Purchase As Purchase)

    Dim WriteRow As Long, TargetSheet As Worksheet, Record As Variant
    
    With Purchase
        Record = Array( _
                    .HoldReason, _
                    .WarehouseNumber, _
                    .VendorName, _
                    .RequireMallCount, _
                    Format(Date, "Mdd"), _
                    .Code, _
                    .RequireQuantity, _
                    .ProductName _
                    )
    End With
    
    Set TargetSheet = Worksheets("�ۗ�")
    WriteRow = TargetSheet.Range("A1").SpecialCells(xlLastCell).Row + 1
    
    TargetSheet.Cells(WriteRow, 4).NumberFormatLocal = "@"
    TargetSheet.Cells(WriteRow, 1).Resize(1, 8).Value = Record
    
End Sub

Private Sub WriteBackupSheet(ByRef Purchase As Purchase)

    Dim WriteRow As Long, TargetSheet As Worksheet, Record As Variant
    
    With Purchase
        Record = Array( _
                    .WarehouseNumber, _
                    .VendorName, _
                    .RequireMallCount, _
                    Format(Date, "Mdd"), _
                    .Code, _
                    .Code, _
                    .ProductName, _
                    .Code, _
                    .PurchaseQuantity _
                    )
    End With
    
    Set TargetSheet = Worksheets("�������i���X�g")
    WriteRow = TargetSheet.Range("A1").SpecialCells(xlLastCell).Row + 1
    
    TargetSheet.Cells(WriteRow, 4).NumberFormatLocal = "@"
    TargetSheet.Cells(WriteRow, 1).Resize(1, 9).Value = Record

End Sub
