Attribute VB_Name = "BuildSheets"
Option Explicit

Sub CreateSorterSheet(Mall As String)

'�P�̏��i�̐U�����p�V�[�g��p��
Worksheets("�U���p�e���v���[�g").Copy after:=Worksheets(Worksheets.Count)
With ActiveSheet
    .Name = Mall & "_�U���p"
    .PageSetup.LeftHeader = Format(Date, "M/dd") & " " & Mall
End With
Dim ForSorterSheet As Worksheet
Set ForSorterSheet = ActiveSheet

'�Z�b�g���i�̐U�����p�V�[�g��p��
Worksheets("�U���p�e���v���[�g").Copy after:=Worksheets(Worksheets.Count)
With ActiveSheet
    .Name = Mall & "_�U���p-�Z�b�g"
    .PageSetup.LeftHeader = Format(Date, "M/dd") & " " & Mall & "-�Z�b�g���i"
End With
Dim ForSorterSetItemSheet As Worksheet
Set ForSorterSetItemSheet = ActiveSheet

'�A�N�e�B�u�ȃV�[�g�̓R�s�[�����V�[�g����󒍃V�[�g�ɕς��Ă���
OrderSheet.Activate

Dim i As Long, k As Long, j As Long, Order(6) As Variant
'�󒍃f�[�^�V�[�g�s�J�E���^
i = 2

'�U�����p�V�[�g�s�J�E���^
j = 2

'�U�����p�Z�b�g�V�[�g�s�J�E���^
k = 2

Do
    '�z��ɍs���i�[
    Order(0) = CStr(Range("D" & i).Value) '6�P�^
    Order(1) = Range("E" & i).Value '���i��
    Order(2) = Range("F" & i).Value '����
    Order(3) = CStr(Range("I" & i).Value) 'JAN
    Order(4) = Range("B" & i).Value '���͂��於
    Order(5) = Range("Q" & i).Value '���݌�
    
    '�]�L�攻��
    '7777�n�܂�Z�b�g�ƃZ�b�g�\�����i�A�󒍎��R�[�h7777***
    If Range("C" & i) Like "7777*" Then
       
        With ForSorterSetItemSheet
            
            .Range("A" & j & ":F" & j).NumberFormatLocal = "@"
            .Range("A" & j & ":F" & j) = Order
            
            '���ʁA���݌ɂ͉E��
            .Range("C" & j).HorizontalAlignment = xlRight
            .Range("F" & j).HorizontalAlignment = xlRight
        
        End With
        
        j = j + 1
          
    Else
    
        With ForSorterSheet
        
            .Range("A" & k & ":F" & k).NumberFormatLocal = "@"
            .Range("A" & k & ":F" & k) = Order
       
           '���ʁA���݌ɂ͉E��
            .Range("C" & k).HorizontalAlignment = xlRight
            .Range("F" & k).HorizontalAlignment = xlRight
       
            '�I�ԂȂ��́A�s�ɐF��t����B
            If OrderSheet.Range("H" & i).Value = "" Then
                     
                With .Range("A" & k & ":F" & k).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent2
                    .TintAndShade = 0.799981688894314
                    .PatternTintAndShade = 0
                End With
             
            End If
        
        End With
        
        k = k + 1
    
    End If

Continue:
    
    i = i + 1

Loop Until IsEmpty(Range("A" & i))


Call Sort.�U���p�V�[�g_�\�[�g(ForSorterSheet)

ForSorterSheet.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
ForSorterSetItemSheet.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous


'�I������ Sheet���e�m��

'�O�̂��ߕ����Ďw��
Call AdjustWidth(ForSorterSheet)
Call AdjustWidth(ForSorterSetItemSheet)

ForSorterSheet.Protect
ForSorterSetItemSheet.Protect

End Sub

Sub OutputPickingData(MallName As String)

'��o�p�t�@�C����p��
'100��/200�ԒI�L�� -2-3�A�d�Z����o
Dim ForSlimsBook As Workbook, ForSlimsSheet As Worksheet
Set ForSlimsBook = PreparePickingBook(MallName & "P�V�[�g" & Format(Date, "MMdd") & "-2-3")
Set ForSlimsSheet = ForSlimsBook.Worksheets(1)

'�o�^�����A�I���� -a
Dim NoEntryItemBook As Workbook, NoEntryItemSheet As Worksheet
Set NoEntryItemBook = PreparePickingBook(MallName & "P�V�[�g" & Format(Date, "MMdd") & "-a")
Set NoEntryItemSheet = NoEntryItemBook.Worksheets(1)

OrderSheet.Activate

Dim i As Long, k As Long, j As Long, Order(6) As Variant
'�󒍃f�[�^�V�[�g�s�J�E���^
i = 2

'�I�����V�[�g�s�J�E���^
j = 2

'100�ԃV�[�g�s�J�E���^
k = 2

'�p�ӂ����u�b�N��1�s���R�s�[
Do
    '�󒍎��R�[�h��7777�͓d�Z��o�f�[�^�Ɋ܂߂Ȃ��B
    If Range("D" & i).Value Like "7777*" Then GoTo Continue

    '��o����R�[�h�̐U��
    'SLIMS�ɓ�������̂̓��P�[�V�����L���6�P�^�̂�
    Dim OrderedCode As String, AddinResultCode As String, Code As String
    
    OrderedCode = CStr(Range("C" & i).Value)
    AddinResultCode = CStr(Range("J" & i).Value)
    
    If AddinResultCode = "" Then
        Code = OrderedCode
    Else
        Code = AddinResultCode
    End If
    
    '�z��ɒ�o�t�@�C��1�s���̃f�[�^���i�[
    
    Order(0) = Range("A" & i).Value '�����ԍ�
    Order(1) = Code '���i�R�[�h
    Order(2) = Range("E" & i).Value '���i��
    Order(3) = Range("F" & i).Value '����
    Order(4) = Range("G" & i).Value '�̔����i
    Order(5) = Range("G" & i).Value '���݌�
    Order(6) = Range("H" & i).Value '�L�����P�[�V����
    
    
    '�]�L�攻��  �R�[�h�������͏����F������Ƃ��āA�擪�[�����J�b�g����Ȃ��悤��
    
    '���P�[�V�����Ȃ�
    If Order(6) = "" Then
        
        NoEntryItemSheet.Range("C" & j).NumberFormatLocal = "@"
        NoEntryItemSheet.Range("B" & j & ":G" & j) = Order
    
        j = j + 1
    
    Else

        ForSlimsSheet.Range("C" & k).NumberFormatLocal = "@"
        ForSlimsSheet.Range("B" & k & ":H" & k) = Order
       
        k = k + 1
    
    End If

Continue:
    
    i = i + 1

Loop Until IsEmpty(Range("A" & i))

'�r��������
ForSlimsSheet.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
NoEntryItemSheet.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous

'�u�b�N��ۑ�
ForSlimsBook.Close SaveChanges:=True
NoEntryItemBook.Close SaveChanges:=True

End Sub

Private Sub AdjustWidth(TargetSheet As Worksheet)
'�� �������ɃA���[�g���o��̂�}�~
Application.DisplayAlerts = False

Dim WidthArr As Variant
WidthArr = Array(14.75, 84.13, 4.25, 15.88, 14.88, 6.63)

TargetSheet.Activate

Dim k As Long
For k = 0 To 5
    TargetSheet.Columns(k + 1).ColumnWidth = WidthArr(k)
Next

Application.DisplayAlerts = True

End Sub

Private Function PreparePickingBook(BookName As String) As Workbook
'�u�b�N����ς��邽�߂ɁA����̏ꏊ�֐�Ƀf�[�^�Ȃ��ŕۑ�����

ThisWorkbook.Worksheets("�s�b�L���O�V�[�g��o�p�e���v���[�g").Copy

ActiveSheet.Name = BookName

'�t�@�C���ۑ�����
'�[���I��Try-Catch�Ńt�@�C����ۑ�����
On Error Resume Next
    
    'Try �ۑ�
  
    ActiveWorkbook.SaveAs FileName:="\\Server02\���i��\�l�b�g�̔��֘A\�s�b�L���O\" & BookName & ".xlsx"

    'catch
    If Err Then
        Err.Clear
        MsgBox "�l�b�g�̔��֘A�Ɍq����Ȃ����߁A" & BookName & "�́A�f�X�N�g�b�v�ɕۑ����܂��B"
        ActiveWorkbook.SaveAs FileName:="C:" & Environ("HOMEPATH") & "\Desktop\" & BookName & ".xlsx"
    End If
    
    'catch2
    If Err Then
        Err.Clear
        MsgBox "�t�@�C����ۑ��ł��܂���ł����B�蓮�ŕۑ����Ă��������B"
    End If

Set PreparePickingBook = ActiveWorkbook

ThisWorkbook.ActiveSheet.Activate

End Function
