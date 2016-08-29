Attribute VB_Name = "BuildSheets"
Option Explicit

Sub �V�[�g�쐬()

OrderSheet.Activate

'InhouseData.TransferOrderSheet

BuildSheets.TransferPickingData
BuildSheets.TransferSorterSheet

End Sub

Private Sub TransferSorterSheet()

Worksheets("�U�����p�ꗗ�V�[�g").PageSetup.LeftHeader = Format(Date, "M/dd") & " Yahoo!�V���b�s���O"
Worksheets("�U�����p�ꗗ�V�[�g-�Z�b�g").PageSetup.LeftHeader = Format(Date, "M/dd") & " Yahoo!�V���b�s���O �Z�b�g"

Dim i As Long, k As Long, j As Long, Order(8) As Variant
'�󒍃f�[�^�V�[�g�s�J�E���^
i = 2

'�l�b�g�p�݌ɃV�[�g�s�J�E���^
j = 2

'�q��1�V�[�g�s�J�E���^
k = 2

Do
    '�z��ɍs���i�[
    Order(0) = CStr(Range("A" & i).Value) '�����ԍ�
    Order(1) = Range("B" & i).Value '���͂��於
    Order(2) = Range("D" & i).Value '�Г��R�[�h
    Order(3) = Range("E" & i).Value '���i��
    Order(4) = Range("F" & i).Value '����
    Order(5) = Range("L" & i).Value 'JAN
    Order(6) = Range("I" & i).Value '���݌�
    Order(7) = Range("J" & i).Value '���P�[�V����
    Order(8) = Range("K" & i).Value '���l



    '�]�L�攻��
       
    If Order(2) Like "7777*" Or Range("C" & i).Value = "Set" Then
       
       Worksheets("�U�����p�ꗗ�V�[�g-�Z�b�g").Range("A" & j & ":I" & j).NumberFormatLocal = "@"
       'Worksheets("�U�����p�ꗗ�V�[�g-�Z�b�g").Range("E" & j & ",G" & j).NumberFormatLocal = "G/�W��"
       Worksheets("�U�����p�ꗗ�V�[�g-�Z�b�g").Range("A" & j & ":I" & j) = Order
    
       j = j + 1
          
    '����ȊO
    Else
    
       Worksheets("�U�����p�ꗗ�V�[�g").Range("A" & k & ":I" & k).NumberFormatLocal = "@"
       Worksheets("�U�����p�ꗗ�V�[�g").Range("A" & k & ":I" & k) = Order
       
       k = k + 1
    
    End If

Continue:
    
    i = i + 1

Loop Until IsEmpty(Range("A" & i))

End Sub

Private Sub TransferPickingData()

Dim i As Long, k As Long, j As Long, Order(4) As Variant
'�󒍃f�[�^�V�[�g�s�J�E���^
i = 2

'�l�b�g�p�݌ɃV�[�g�s�J�E���^
j = 2

'�q��1�V�[�g�s�J�E���^
k = 2

Do
    '�z��ɍs���i�[
    Order(0) = Range("D" & i).Value '�Г��R�[�h
    Order(1) = Range("E" & i).Value '���i��
    Order(2) = Range("F" & i).Value '����
    Order(3) = Range("G" & i).Value '���t�[�̔����i
    Order(4) = Range("H" & i).Value '����

    '�]�L�攻��[
    '���P�[�V�����Ȃ�
    If Range("J" & i).Value = "" Then
        
        If Not Order(0) Like "7777*" Then
           
           Worksheets("�l�b�g�p�݌�").Range("B" & j).NumberFormatLocal = "@"
           Worksheets("�l�b�g�p�݌�").Range("B" & j & ":F" & j) = Order
        
           j = j + 1
        
        End If
        
    '���P�[�V�����L��
    Else
    
       Worksheets("�q��1").Range("B" & k).NumberFormatLocal = "@"
       Worksheets("�q��1").Range("B" & k & ":F" & k) = Order
       
       k = k + 1
    
    End If

Continue:
    
    i = i + 1

Loop Until IsEmpty(Range("A" & i))

End Sub
