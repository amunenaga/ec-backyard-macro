Attribute VB_Name = "FetchWebSyokon"
Option Explicit

'VBA��IE���g����WEB�̃f�[�^���擾������@
'http://excel-ubara.com/excelvba5/EXCELVBA222.html
'http://www.vba-ie.net/ieobject/refresh.html

Type PurchaseLog
    
    Code As String
    PurchaseDate As Date
    WarehouseNum As Integer
    PurchaseQuantity As Long
    NonArrivalQty As Long
    Po As Long
    LastArrival As Date

End Type

Sub CheckNonArrival()
'WEB��������A���c���Ȃ������ׂĖ����ׂ��L��Δ��l��֒ǋL

'�擾�ł��Ȃ��Ă������͂Ƃɂ������s
'On Error Resume Next

Dim CodeRange As Range, r As Range
Set CodeRange = Range(Cells(2, 7), Cells(2, 7).End(xlDown))

For Each r In CodeRange
    
    Dim Code As String
    
    Code = r.Value
    If Len(Code) > 6 Then GoTo Continue
    
    Dim LastPur As PurchaseLog, InitPur As PurchaseLog
    LastPur = InitPur
    LastPur = FetchRecentPurchase(Code)
    
    If LastPur.NonArrivalQty > 0 Then
        Cells(r.Row, 2).Value = Cells(r.Row, 2).Value & IIf(Cells(r.Row, 2).Value = "", "", " ") & "������" & LastPur.NonArrivalQty & "�� " & Format(LastPur.PurchaseDate, "M��d��") & "��z��"
    End If

Continue:

Next

End Sub

Private Function FetchRecentPurchase(ByVal Code As String) As PurchaseLog
'WEB��������HTML�o�R�Œ��߂̎�z�󋵂�1���擾����

    Dim CurrentCode As PurchaseLog
    CurrentCode.Code = Code

    On Error Resume Next
            
        Dim SyokonPage As InternetExplorerMedium
        Set SyokonPage = New InternetExplorerMedium
    
        SyokonPage.Navigate "http://server02/gyomu/SK_IZoom.asp?ICode=" & Code & "&C5="
        Call untilReady(SyokonPage)
        
        '�I�u�W�F�N�g�ϐ���DOM
        Dim DivPurchaseLog As Object
        Set DivPurchaseLog = SyokonPage.Document.getElementsByName("t1") '�ŋ߂̔����󋵂Ɠ��׈ē� Div�^�O��ID��t1
        Dim RecentRow As Object
    
        Set RecentRow = DivPurchaseLog(0).all.Item(13) '�ŋ߂̔����󋵂Ɠ��׈ē��e�[�u��2�s�ڂ�DOM
        
        With CurrentCode
            .PurchaseDate = CDate(RecentRow.all.Item(0).innertext)
            .WarehouseNum = RecentRow.all.Item(1).innertext
            .PurchaseQuantity = RecentRow.all.Item(2).innertext
            .NonArrivalQty = IIf(RecentRow.all.Item(3).innertext = "����", 0, RecentRow.all.Item(3).innertext)
            .Po = RecentRow.all.Item(4).innertext
            If Not RecentRow.all.Item(5).innertext Like "*-*" Then
                .LastArrival = CDate(RecentRow.all.Item(5).innertext)
            End If
        End With
        
        SyokonPage.Quit

    On Error GoTo 0
    
    FetchRecentPurchase = CurrentCode

End Function

Private Sub untilReady(objIE As Object, Optional ByVal WaitTime As Integer = 20)
'WEB�����̃��X�|���X�҂��̂��߂̃v���V�[�W��

    '�T�[�o�[���X�|���X�ҋ@
    Dim starttime As Date
    starttime = Now()
    Do While objIE.Busy = True Or objIE.ReadyState <> READYSTATE_COMPLETE
        DoEvents
        If Now() > DateAdd("S", WaitTime, starttime) Then
            Exit Do
        End If
    Loop
    
    '���[�f�B���O��ʂ̕\����ɁA�ڍ׃f�[�^�����I�ɍĕ`�悳���̂�1�b�ҋ@
    Dim WaitEnd As Date
    WaitEnd = DateAdd("S", 2, Now())
    Do While Now() < WaitEnd
        DoEvents
    Loop
    
    DoEvents

End Sub
