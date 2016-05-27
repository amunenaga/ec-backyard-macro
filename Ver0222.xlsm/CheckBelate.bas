Attribute VB_Name = "CheckBelate"
Public Sub �x���`�F�b�N()

Dim BelateList As Dictionary
Set BelateList = MakeBelateList()

For Each v In BelateList  '�x�����X�g��MsgBox�ŕ\�����邽��String�ŏo�͂��ĘA���B
    
Dim IdList As String
    
    IdList = IdList & vbLf & Format(BelateList(v).OrderDate, "MM/dd") & "  No." & BelateList(v).Id

Next

Application.ScreenUpdating = True

If BelateList.Count > 0 Then

    MsgBox prompt:="������/���A����3���ȏ�o�߂��Ă��钍��" & vbLf _
            & IdList & vbLf _
            & vbLf _
            & BelateList.Count & "������܂��B", _
            Buttons:=vbExclamation, _
            Title:="�`�F�b�N����"

Else

    MsgBox "������/���A����3���o�߂��Ă��钍���͂���܂���B", _
            Buttons:=vbInformation, _
            Title:="�`�F�b�N����"

End If

End Sub

Public Sub �x�����X�g�o��()

Dim BelateList As Dictionary
Set BelateList = MakeBelateList()

Workbooks.Add

'�V�K�ǉ��t�@�C���Ƀw�b�_�[�쐬
With ActiveSheet
    
    .Range("A1").Value = "�o�׏󋵊m�F " & Format(Date, "m��d��")
    .Range("A2:I2") = Array("�󒍓�", "�����ԍ�", "�����Җ�", "Line", "���i�R�[�h", "���i��", "����", "�o�׏�", "�o�ד�", "�����ԍ�")

End With

Dim i As Integer
i = 3


For Each v In BelateList
  
    Id = BelateList(v).Id
    
    '�����ԍ�����A���c�ꗗ�V�[�g�̍s�ԍ������A��������z��Ɋi�[
    With ThisWorkbook.Worksheets("���c�ꗗ")
        Dim r As Long, rng As Range
        r = .Range("B:B").Find(Id).Row
        Dim arr
        arr = .Range("A" & r & ":" & "G" & r)
    
    End With
    
    '�쐬�����V�K�u�b�N�ɓ\��t���čs��
    ActiveSheet.Range(Cells(i, 1), Cells(i, 7)) = arr
    
    i = i + 1

Next

Debug.Print i
'Line�ԍ�����폜
ActiveSheet.Columns("d:d").Delete

ActiveSheet.Range(Cells(3, 1), Cells(1, 1).End(xlDown)).NumberFormatLocal = "m""��""d""��"";@"

End Sub

Private Function MakeBelateList() As Dictionary
'belated arrival�ŉ����̂��Ƥ�x����Belate�œ��ꂵ�܂��B
'OrderList���쐬���āABelate=�x���`�F�b�N�����āA�Y��������BelateList�ɒǉ����܂��B

'���c�ꗗ��OrderSheet

OrderSheet.Activate

Application.ScreenUpdating = False

'���������i��OrderList�����܂�
Dim UndispatchList As Dictionary
Set UndispatchList = OrderSheet.getUndispatchOrders

Dim o As order
Dim v As Variant

Dim BelateList As Dictionary
Set BelateList = New Dictionary

For Each v In UndispatchList 'OrderList�̌X��Order�ɂ��āA�`�F�b�N
    Set o = UndispatchList(v)
      
    If checkBelateDispatch(o) Then 'checkBelateDispatch Function�Ń`�F�b�N
              
        'AlertPiriod���SendMailDate�����=Purchase��Ɉ�x�A�����Ă��钍���́A�x�����X�g�ɉ����Ȃ�
        'AlertPiriod��EstimatedArrivalDate=���ד����w���Ă���΁A�ʏ탋�[�e�B���ł͓��ד��͘A���������ɂȂ邽�߁ADateDiff�Ő��̒l�ɂȂ�B
        'Day�ł�����������Ȃ��̂ŁA1�����𒴂���ƃA���[�g�オ��Ȃ����A���ח\�肩��O���o�߂Ƃ��̃X�p���ł̔����R��A�A���R���c���������̂ō\��Ȃ��B
                  
        If DateDiff("d", o.AlertPiriod, o.SendMailDate) < 0 Then
                    BelateList.Add o.Id, o
        
        End If
    
    End If
    
Next

Set MakeBelateList = BelateList

End Function

Private Function checkBelateDispatch(order As order) As Boolean

'�P�̖̂������`�F�b�J�[

    '�U���̏ꍇ�͓����A���ςŁA7���𒴂��Ă���Βx���B���t�[�̎��������Œ��������14����Ƀ|�C���g�����m�肷��̂�
    If order.IsWaitingPayment And DateDiff("d", order.AlertPiriod, Date) > 7 Then
        TmpBl = True
        GoTo re
    Else
        TmpBl = False
    End If

    'Order�I�u�W�F�N�g�̃A���[�g�N�Z���Ɩ{���̍���2�𒴂��遁�O���ȏ��True

    If DateDiff("d", order.AlertPiriod, Date) > 2 Then
        TmpBl = True
    Else
        TmpBl = False
    End If

re:
checkBelateDispatch = TmpBl

End Function
