Attribute VB_Name = "FixYahataPcs"
Option Explicit

Sub ExtractProcuctName()

'���K�\���I�u�W�F�N�g��RegExp
Dim re As Object
Set re = CreateObject("VBScript.RegExp")

re.Global = True

'�����茟�o�p�^�[��
re.Pattern = "[0-9]{1,4}����"

Dim i As Long
For i = 131070 To 135086
    
    If Not Cells(i, 1).Value Like "*-*" Then GoTo Continue
    
    Dim str As String
    str = Cells(i, 2).Value
    Cells(i, 2).Value = re.Replace(str, "") & Cells(i, 3).Value
    
Continue:

Next

End Sub

Sub ExtractYahataJanPcs()

'���K�\���I�u�W�F�N�g��RegExp
Dim re As Object
Set re = CreateObject("VBScript.RegExp")

re.Global = True

'�����茟�o�p�^�[��
re.Pattern = "[0-9]{1,4}����"

Dim i As Long
For i = 131070 To 135086
    
    If Cells(i, 1).Value Like "*-*" Then GoTo Continue
    Dim str As String
    
    str = Cells(i, 2).Value
    
    Dim Matches
    
    Set Matches = re.Execute(str)
    
    If Matches.Count = 0 Then GoTo Continue
    Cells(i, 3).Value = Matches(0)

Continue:

Next

End Sub

Sub ReCalcurationSetQty(RowCount As Long)

Do
    Dim Code As String
    Code = Cells(RowCount, 1).Value
    
    Dim SinglexQty As String, SinglexJAN As String, CurrentSinglexJan As String
    
    '�n�C�t���L��̎��͕������čĎZ�o
    If Code Like "*-*" Then
             
        Dim SeparatedCode As Variant
        SeparatedCode = Split(Code, "-", 2)
        
        Dim SetQty As String
        SetQty = CStr(Val(SeparatedCode(1)))
        
        Dim SetQtyStr As String
        SetQtyStr = SinglexQty & "�~" & SetQty
       
        '�Z�o�����������\�L��C��֊i�[
        Cells(RowCount, 3).Value = SetQtyStr
        
        CurrentSinglexJan = SeparatedCode(0)
        
    Else
    '�n�C�t���Ȃ����P��JAN�̏ꍇ�A�P�̂ł̓��e�ʂ��i�[
    
        SinglexQty = Cells(RowCount, 3).Value
        SinglexJAN = Cells(RowCount, 1).Value
        
        CurrentSinglexJan = Cells(RowCount, 1).Value
    
    End If
    
    '�Q�Ɠn���Ȃ̂Ō��J�E���^��i�߂Ă���
    RowCount = RowCount + 1

Loop While CurrentSinglexJan = SinglexJAN

End Sub

Sub Recalc()

Dim i As Long
For i = 131070 To 140000

    Call ReCalcurationSetQty(i)

Next

End Sub
