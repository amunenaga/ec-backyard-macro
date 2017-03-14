Attribute VB_Name = "BuildSheets"
Option Explicit
Sub CreateSorterSheet(MallName As String)

'�P�̏��i�̐U�����p�V�[�g��p��
Worksheets("�U���p�e���v���[�g").Copy after:=Worksheets(Worksheets.Count)
With ActiveSheet
    .Name = MallName & "_�U���p"
    .PageSetup.RightHeader = Format(Date, "M/dd") & " " & MallName
End With
Dim ForSorterSheet As Worksheet
Set ForSorterSheet = ActiveSheet

'�Z�b�g���i�̐U�����p�V�[�g��p��
Worksheets("�U���p�e���v���[�g").Copy after:=Worksheets(Worksheets.Count)
With ActiveSheet
    .Name = MallName & "_�U���p-�Z�b�g"
    .PageSetup.RightHeader = Format(Date, "M/dd") & " " & MallName & "-�Z�b�g���i"
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

    '�����œn���ꂽ���[���ȊO�͔�΂�
    If Not Range("F" & i).Value Like (MallName & "*") Then GoTo Continue

    '�z��ɍs���i�[
    Order(0) = CStr(Range("B" & i).Value) '�󒍎��R�[�h
    Order(1) = Range("C" & i).Value '���i��
    Order(2) = Range("D" & i).Value '�󒍐���
    Order(3) = CStr(Range("L" & i).Value) 'JAN
    Order(4) = Range("G" & i).Value '���͂��於
    Order(5) = Range("N" & i).Value '���݌�
    
    
    '���݌ɂ��擾�ł��ĂȂ��Ƃ��́A������C�A�E�g�̊֌W�̂��ߋ�1��������Ă���
    If Order(5) = "" Then Order(5) = " "
    
    '�]�L�攻��
    '7777�n�܂�Z�b�g�ƃZ�b�g�\�����i�A�󒍎��R�[�h7777***
    If Range("B" & i) Like "7777*" Then
       
       Order(0) = Range("I" & i).Value
       
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
       
           '���ʁA���݌ɂ͉E�񂹁AJAN�͒���
            .Range("C" & k).HorizontalAlignment = xlRight
            .Range("D" & k).HorizontalAlignment = xlCenter
            .Range("F" & k).HorizontalAlignment = xlRight
       
            '�I�ԂȂ��́A�s�ɐF��t����B
            If OrderSheet.Range("K" & i).Value = "" Then
                     
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

    '�����œn���ꂽ���[���ȊO�͔�΂�
    If Not Range("G" & i).Value Like (MallName & "*") Then GoTo Continue
    
    '�󒍎��R�[�h��7777�͓d�Z��o�f�[�^�Ɋ܂߂Ȃ��B
    If Range("I" & i).Value Like "7777*" Then GoTo Continue

    '��o����R�[�h�̐U��
    'SLIMS�ɓ�������̂̓��P�[�V�����L���6�P�^�̂�
    Dim OrderedCode As String, AddinResultCode As String, Code As String
    
    OrderedCode = CStr(Range("B" & i).Value)
    AddinResultCode = CStr(Range("M" & i).Value)
    
    If AddinResultCode = "" Then
        Code = OrderedCode
    Else
        Code = AddinResultCode
    End If
    
    '�z��ɒ�o�t�@�C��1�s���̃f�[�^���i�[
    
    Order(0) = CStr(Range("A" & i).Value) '�����ԍ�
    Order(1) = CStr(Code) '���i�R�[�h
    Order(2) = Range("C" & i).Value '���i��
    Order(3) = Range("E" & i).Value '����
    Order(4) = Range("D" & i).Value '�̔����i
    Order(5) = Range("N" & i).Value '���݌�
    Order(6) = Range("K" & i).Value '�L�����P�[�V����
    
    '�]�L�攻��  �R�[�h�������͏����F������Ƃ��āA�擪�[�����J�b�g����Ȃ��悤��
    
    '���P�[�V�����Ȃ�
    If Order(6) = "" Then
        
        NoEntryItemSheet.Range("B" & j & ":C" & j).NumberFormatLocal = "@"
        NoEntryItemSheet.Range("B" & j & ":H" & j) = Order
    
        j = j + 1
    
    Else

        ForSlimsSheet.Range("B" & k & ":C" & k).NumberFormatLocal = "@"
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

Private Function PreparePickingBook(BookName As String) As Workbook
'�u�b�N����ς��邽�߂ɁA����̏ꏊ�֐�Ƀf�[�^�Ȃ��ŕۑ�����

'�����̖��O�ŐV�K�u�b�N���쐬����
ThisWorkbook.Worksheets("�s�b�L���O�V�[�g��o�p�e���v���[�g").Copy
ActiveSheet.Name = BookName

'��U�V�K�쐬�u�b�N��ۑ����邱�ƂŃu�b�N����ύX����
'�V�K�쐬�t�@�C���̕ۑ����̓t�@�C���t�H�[�}�b�g�𖾎�����K�v�Ȗ͗l
    Dim SavePath As String
    Const PICKING_FOLDER As String = "\\Server02\���i��\�l�b�g�̔��֘A\�s�b�L���O\"
    
    If Dir(PICKING_FOLDER, vbDirectory) <> "" Then
        '���ɖ{���t�@�C��������΁A�����t���ĕۑ�
        If Dir(PICKING_FOLDER & BookName & ".xlsx") = "" Then
            SavePath = PICKING_FOLDER & BookName
        Else
            SavePath = PICKING_FOLDER & Format(Time, "hhmm") & BookName
        End If
        
            ActiveWorkbook.SaveAs Filename:=SavePath, FileFormat:=xlWorkbookDefault
    
    Else
        
        Dim DeskTopPath As String
        If Dir(DeskTopPath & BookName & ".xlsx") = "" Then
            DeskTopPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\" & BookName
        Else
            DeskTopPath = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\" & Format(Time, "hhmm") & BookName
        End If
        
        MsgBox "�l�b�g�̔��֘A�Ɍq����Ȃ����߁A" & BookName & "���f�X�N�g�b�v�ɕۑ����܂��B"
            
        ActiveWorkbook.SaveAs Filename:=DeskTopPath, FileFormat:=xlWorkbookDefault
    
    End If

Set PreparePickingBook = ActiveWorkbook

ThisWorkbook.ActiveSheet.Activate

End Function

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
