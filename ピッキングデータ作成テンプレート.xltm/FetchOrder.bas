Attribute VB_Name = "FetchOrder"
Option Explicit

Sub ReadMeisai(Path As String)

'Meisai.CSV��OrderSheet=�����ꗗ�ɒǋL����

Dim FSO As Object
Set FSO = New FileSystemObject

Dim TS As Textstream
Set TS = FSO.OpenTextFile(Path, ForReading)

Dim i As Long
i = 1 '���ڍs�͏o�͂��Ȃ��̂ŁAi��1�s�ڂ���J�n
    
Do Until TS.AtEndOfStream
    
' �s�����o���ĕK�v�ȍ��ڂ݂̂�z��ɓ��꒼��
    Dim LineBuf As Variant
    LineBuf = Split(TS.ReadLine, """,""")
       
    Dim j As Integer
    For j = 0 To UBound(LineBuf)
        LineBuf(j) = Trim(Replace(LineBuf(j), Chr(34), "")) 'chr(34)�� " [���p��d���p��]
    
    Next
    
    'CSV���w�b�_�[ 0:Order ID/1:Line ID/2:Quantity/3:Product Code/4:Description/5:Option Name/6:Option Value/7:Unit Price/
    '���[�v���ڂł̓w�b�_�[�Ȃ̂ŁA�C���N�������g�֔��
    
    If LineBuf(0) = "Order ID" Then GoTo increment
        
    ':ToDo ��������V�[�g�A�Z���̒l�̂�����̂ŕ�����������������������Ȃ��B
    
    With Worksheets("�󒍃f�[�^�V�[�g")
        'A��A�����ԍ�
        .Range("A" & i).Value = LineBuf(0)
        
        'C��A�󒍎��̏��i�R�[�h
        .Range("C" & i).NumberFormatLocal = "@"
        .Range("C" & i).Value = LineBuf(3)
        
        'D��A�A�h�C�����s�p��6�P�^�������R�[�h�A��������JAN
        '�󗓂����肦��̂ŁA�s�b�L���O�f�[�^�E�U�����X�g�ɓ]�L���ɋ󗓔��肷��
        
        '6�P�^�Ȃ炻�̂܂ܓ����
        If LineBuf(3) Like "######" Then
            .Range("D" & i).NumberFormatLocal = "@"
            .Range("D" & i).Value = LineBuf(3)
        
        '����5�P�^�͓��Ƀ[����ǋL
        ElseIf LineBuf(3) Like "#####" Then
            
            .Range("D" & i).NumberFormatLocal = "@"
            .Range("D" & i).Value = "0" & LineBuf(3)
        
        'JAN�����̂܂ܓ����
        ElseIf LineBuf(3) Like String(13, "#") Then
            
            .Range("D" & i).NumberFormatLocal = "@"
            .Range("D" & i).Value = LineBuf(3)
        
        End If
        
        'E��F���i��  F��F�󒍐���  G��F����
        .Range("E" & i).Value = LineBuf(4)
        .Range("F" & i).Value = LineBuf(2)
        .Range("G" & i).Value = LineBuf(7)
        
        'CSV1�s�����[�h����
        

        '�Z�b�g���� 7777�n�܂�
        If .Range("C" & i).Value Like "7777*" Then
            
            Call SetParser.ParseItems(.Range("D" & i))
            
            'ParseItems�ōs���}�������̂ŁA�s�J�E���^���Z�b�g������
            i = OrderSheet.Range("A1").CurrentRegion.Rows.Count
            
        
        End If
    
        '�P�́��Z�b�g���� �n�C�t���܂ރR�[�h�Ȃ番�������֓�����
        
        If .Range("C" & i).Value Like "*-*" Then
        
            Call SetParser.ParseScalingSet(.Range("C" & i))
        
        End If
    
    End With
    
increment:
    i = i + 1

Loop

TS.Close

SetParser.CloseSetMasterBook

End Sub

Sub ReadTyumonH(Path As String)

Dim FSO As Object
Set FSO = New FileSystemObject

Dim TS As Textstream
Set TS = FSO.OpenTextFile(Path, ForReading)

'�Ǎ��ϒ����ԍ��̃����W���Z�b�g�AA1����A��̔ԍ�����ŏI�Z���܂�
Dim LoadedOrderRange As Range
Set LoadedOrderRange = OrderSheet.Cells(1, 1).Resize(OrderSheet.Cells(1, 1).SpecialCells(xlCellTypeLastCell).Row, 1)

Do Until TS.AtEndOfStream
    
' �s�����o���ĕK�v�ȍ��ڂ݂̂�z��ɓ��꒼��
    Dim LineBuf As Variant
    LineBuf = Split(TS.ReadLine, """,""")
    
    '0=1���=�����ԍ��A�����Җ��A�v�]�A���ϕ��@�A�N�[�|���l����
    Dim Order As Variant
    Order = Array(LineBuf(0), LineBuf(5), LineBuf(36), LineBuf(34), LineBuf(43))
        
    Dim j As Integer
    For j = 0 To UBound(Order)
        Order(j) = Trim(Replace(Order(j), Chr(34), "")) 'chr(34)�� " [���p��d���p��]
    
    Next

    '�����ԍ��̍s�𒲂ׂ�
    '�����ԍ���Dobule�^�œ����Ă���BCSV��String�^�AMatch�֐��̕Ԓl��Double�^
    
    Dim FindRow As Double
    
    On Error Resume Next
        
        FindRow = WorksheetFunction.Match(CDbl(Order(0)), LoadedOrderRange, 0)
        
        If Err Then
            Err.Clear
            GoTo Continue
        End If
    
    On Error GoTo 0
        
    Dim i As Long
    i = 0
    
    '�����Җ����L�� �I�t�Z�b�g���A�Y�������ԍ��̑S�Ă̍s�֋L��
    Do While Range("A" & FindRow).Offset(i, 0).Value = CDbl(Order(0))
        
        Range("A" & FindRow).Offset(i, 1).Value = LineBuf(5)
        i = i + 1
    
    Loop

        
Continue:
    
Loop

TS.Close

End Sub
