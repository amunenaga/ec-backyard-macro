Attribute VB_Name = "FetchOrder"
Option Explicit

Type RawOrder

    Serial As String        '�N���X���[���ō̔Ԃ���A��
    
    MallId As String        '�󒍃��[���R�[�h  01�y�V 02���t�[ 03Amazon
    MallName As String      '�󒍃��[������
    OrderId As String       '�e���[���̎󒍔ԍ�
    
    Addressee As String     '����於
    
    Code As String          '�󒍎��̏��i�R�[�h
    ProductName As String   '���[���f�ڂ̏��i��
    Quantity As String      '�󒍐���
    
    Price As String         '�󒍋��z

End Type

Sub ReadClossMallCsv(Path As String)

'�N���X���[���̎�CSV���󒍃f�[�^�V�[�g�ɓǂݍ���

Dim FSO As Object
Set FSO = New FileSystemObject

Dim TS As Textstream
Set TS = FSO.OpenTextFile(Path, ForReading)

Dim WriteRow As Long
WriteRow = 1 '���ڍs�͏o�͂��Ȃ��̂ŁAi��1�s�ڂ���J�n
    
Do Until TS.AtEndOfStream
    
' �s�����o���ĕK�v�ȍ��ڂ݂̂�z��ɓ��꒼��
    Dim Col As Variant
    Col = Split(TS.ReadLine, ",")
    
    '���[�v���ڂł̓w�b�_�[�Ȃ̂ŁAContinue
    If Col(0) = "�Ǘ��ԍ�" Then GoTo Continue
    
    Dim Order As RawOrder
    
    With Order
        .Serial = Col(0)
        
        .Code = Col(1)
        .ProductName = Col(2)
        .Quantity = Col(3)
        .Price = Col(4)
                
        .MallName = Col(8)
        .Addressee = Col(10)
        .OrderId = Col(13)

    End With
    
    Call WriteSheet(Order)

    '�ŏI�s����肵�ăZ�b�g����
    Dim CurrentCodeCell As Range
    Set CurrentCodeCell = Cells(Range("A1").SpecialCells(xlCellTypeLastCell).Row, 2)
    
    
    '7777�n�܂�Z�b�g����
    If CurrentCodeCell.Value Like "7777*" Then

        Call SetParser.ParseItems(CurrentCodeCell)
    
    End If

    '�P�́��Z�b�g����
    If CurrentCodeCell.Value Like "*-*" Then
    
        Call SetParser.ParseScalingSet(CurrentCodeCell)
    
    End If

Continue:

Loop

TS.Close

SetParser.CloseSetMasterBook

End Sub

Sub WriteSheet(ByRef Order As RawOrder)
'�����f�[�^�̔z����󂯎���āA�ŏI�s�̒����֒ǋL
    With Worksheets("�󒍃f�[�^�V�[�g")
    
        Dim WriteRow As Long
        WriteRow = .Range("A1").SpecialCells(xlCellTypeLastCell).Row + 1
        
        'A�� �N���X���[���A��
        .Range("A" & WriteRow).NumberFormatLocal = "@"
        .Range("A" & WriteRow).Value = Order.Serial
        
        'B��A�󒍎��̏��i�R�[�h
        .Range("B" & WriteRow).NumberFormatLocal = "@"
        .Range("B" & WriteRow).Value = Order.Code
        
        
        'C��F���i��  D��F����   E��F�󒍐���  F��F�󒍔ԍ� G��F���[����  H��F���͂��於
        .Range("C" & WriteRow).Value = Order.ProductName
        .Range("D" & WriteRow).Value = Order.Price
        .Range("E" & WriteRow).Value = Order.Quantity
        .Range("F" & WriteRow).Value = Order.OrderId
        .Range("G" & WriteRow).Value = Order.MallName
        .Range("H" & WriteRow).Value = Order.Addressee
        
        'I��A�A�h�C�����s�p��6�P�^�������R�[�h�A��������JAN
        '�󗓂����肦��̂ŁA�s�b�L���O�f�[�^�E�U�����X�g�ɓ]�L���ɋ󗓔��肷��
        .Range("I" & WriteRow).NumberFormatLocal = "@"
        
        '6�P�^�Ȃ炻�̂܂ܓ����
        If Order.Code Like String(6, "#") Then
            .Range("I" & WriteRow).Value = Order.Code
        
        '����5�P�^�͓��Ƀ[����ǋL
        ElseIf Order.Code Like String(5, "#") Then
            
            .Range("I" & WriteRow).Value = "0" & Order.Code
        
        'JAN�����̂܂ܓ����
        ElseIf Order.Code Like String(13, "#") Then
            
            .Range("I" & WriteRow).Value = Order.Code
        
        End If
    
        'J�� �K�v���� 6�P�^/JAN�ɑ΂��ĕK�v�Ȑ���
        '�Z�b�g������ɏ�����������̂ŁA��U�󒍐��ʂ�����B
        .Range("J" & WriteRow).Value = Order.Quantity
    
    End With
End Sub

