Attribute VB_Name = "Main"
Option Explicit

Sub �󒍃`�F�b�N���X�g����()

'CSV�Ǎ��A��ƃV�[�g�փR�s�[
Importer.CSV�Ǎ�
Transfer.��ƃV�[�g�փf�[�^���o

'��ƃV�[�g�ł̃f�[�^�C������
Worksheets("��ƃV�[�g").Activate

'��ƃV�[�g�f�[�^�L���`�F�b�N
If Worksheets("��ƃV�[�g").Range("A2").Value = "" Then
    MsgBox Prompt:="�y�V�������܂܂�Ȃ�CSV�f�[�^�ł��A�������I�����܂��B"
    ThisWorkbook.Close SaveChanges:=False
End If

SetParser.�Z�b�g����
Transfer.�X�܎��ʔԍ��U��
Transfer.�Z������
Transfer.JAN�]�L
Transfer.���t�t�H�[�}�b�g�ύX
Transfer.�y�V���i���C��

Transfer.��o�p�V�[�g�֓]�L

'�Z�b�g���i���X�g�u�b�N�����
Dim w As Workbook
For Each w In Workbooks
    If w.Name = "��ď��iؽ�.xls" Then w.Close False
Next

Worksheets("�A�b�v���[�h�V�[�g").Range("A1").Select

'�捞�A��ƃV�[�g���݂̃G�N�Z���t�@�C����ۑ�
Application.DisplayAlerts = False
    
    ActiveWorkbook.SaveAs FileName:=ActiveWorkbook.Path & "\�󒍃`�F�b�N���X�g_" & Format(Date, "yyyymmdd") & ".xlsm", FileFormat:=52
        
Application.DisplayAlerts = True

'�f�[�^�x�[�X�ւ̓o�^�������s
Call CodeI2JAN_E

End Sub


Sub CodeI2JAN_E()
'EC�ϊ���p�i�󒍃`�F�b�N���X�g����.xltm�g���ݐ�p���[�`��
'�Q�Ɛݒ��ADO_Library���K�{�i2.X �n����)


Dim Cnn As New ADODB.Connection
Dim Cmd As New ADODB.Command
Dim Rs As New ADODB.Recordset

Dim SQL_W As String
Dim SQL_W1 As String
Dim R As Long
Dim i As Long
Dim j As Long
Dim IC As Long
Dim JC As String
Dim T_D As Variant
Dim MB_C As Variant
Dim ReceiveNo As Long
Dim Old_Status As String

Old_Status = Application.StatusBar

MB_C = MsgBox("JAN�ϊ���DB�ւ̏������݂��J�n���܂�", vbOKCancel)

If MB_C = vbOK Then

    Cnn.Open "PROVIDER=SQLOLEDB;Server=itoserver3;Database=ITOSQL_REP;UID=sa;PWD=ito;"
    Cmd.CommandTimeout = 180
    Set Cmd.ActiveConnection = Cnn

    Range("A2").Select
    Selection.End(xlDown).Select
    R = ActiveCell.Row

    ReceiveNo = Range("A2").Value
    Application.StatusBar = "�`�[�ԍ��d���`�F�b�N���E�E�E"
    Application.ScreenUpdating = False

        For i = 2 To R
            Cells(i, 1).Select
            If Len(Cells(i, 1).Value) <> 0 Then
                IC = Val(Cells(i, 1).Value)
                SQL_W = "SELECT �󒍔ԍ� From EC_order_base Where �󒍔ԍ� = " & IC & ";"
                Set Rs = Cnn.Execute(SQL_W)
                If Not Rs.EOF Then
                    Cells(i, 15).Value = 1
                End If
            End If
       Next i
    

    SQL_W = "SELECT ���i�R�[�h, JAN�R�[�h FROM ���i�}�X�^ WHERE (���i�R�[�h = "
    
    Application.ScreenUpdating = True
    ActiveWorkbook.RefreshAll

    Application.StatusBar = "�󒍃f�[�^�[��DB�ɏ������ݒ��E�E�E"
    
    Application.ScreenUpdating = False

        For i = 2 To R
            Cells(i, 2).Select
            If Len(Cells(i, 2).Value) <> 0 Then
                IC = Val(Cells(i, 2).Value)
                SQL_W1 = SQL_W & IC & ")"
                Set Rs = Cnn.Execute(SQL_W1)
                If Not Rs.EOF() Then
                    JC = Rs(1)
                    Cells(i, 3).Value = JC
                End If
            End If
        Next i

    Application.ScreenUpdating = True
    ActiveWorkbook.RefreshAll

    Range("A2").Select
    
    Application.ScreenUpdating = False

    For i = 2 To R
        If Cells(i, 15).Value <> 1 Then
            T_D = Cells(i, 1).Value & ","
        
            If Cells(i, 2).Value = "" Then
                T_D = T_D & "Null,"
            Else
                T_D = T_D & Cells(i, 2).Value & ","
            End If
        
            T_D = T_D & "'" & Cells(i, 3).Value & "',"
            T_D = T_D & "'" & Cells(i, 4).Value & "',"
            T_D = T_D & Cells(i, 5).Value & ","
            T_D = T_D & "'" & Cells(i, 6).Value & "',"
            T_D = T_D & "'" & Cells(i, 7).Value & "',"
            T_D = T_D & "'" & Cells(i, 8).Value & "',"
            T_D = T_D & Cells(i, 9).Value & ","
            T_D = T_D & "'" & Cells(i, 10).Value & "',"
            T_D = T_D & "'" & Cells(i, 11).Value & "',"
            T_D = T_D & "'" & Cells(i, 12).Value & "',"
            T_D = T_D & "'" & Cells(i, 13).Value & "',"
            T_D = T_D & Cells(i, 14).Value & ","
            T_D = T_D & "0"
            

            SQL_W = "INSERT INTO EC_Order_Base VALUES (" & T_D & ")"

            Cnn.Execute (SQL_W)

            Cells(i, 15).Value = 1

        End If

    Next i
    
    Application.ScreenUpdating = True
    ActiveWorkbook.RefreshAll

    Application.StatusBar = "�n���f�B�p�f�[�^�[��DB�ɏ������ݒ��E�E�E"
    
    Application.ScreenUpdating = False

    SQL_W = "INSERT INTO EC_order ( �[�i���ԍ�, ���i�R�[�h, �o�[�R�[�h, ���i��, �󒍐���, ���[��, �ڋq��, �t�F�[�Y�ύX����, �����t�F�[�Y, �Z��, �L�����Z��, �󒍃�����, �󒍖��׎}��) "
    SQL_W = SQL_W & "SELECT EC_Order_Base.�󒍔ԍ�, EC_Order_Base.���i�R�[�h, EC_Order_Base.JAN�R�[�h, EC_Order_Base.���i����, EC_Order_Base.����, EC_Order_Base.�[�i���敪, EC_Order_Base.�͂��於��, GETDATE() AS ��1, 1 AS ��2, �͂���Z��, 0, EC_Order_Base.�󒍃�����, EC_Order_Base.�󒍖��׎}�� "
    SQL_W = SQL_W & "FROM EC_Order_Base WHERE EC_Order_Base.�]�� = 0;"

    Cnn.Execute (SQL_W)
    
    SQL_W = "UPDATE EC_Order_Base SET �]�� = 1 WHERE �]�� = 0;"
    Cnn.Execute (SQL_W)
    
    Application.ScreenUpdating = True
    ActiveWorkbook.RefreshAll


    MsgBox "�f�[�^�x�[�X�ւ̏��� �I��", vbInformation

    Cnn.Close

Else

MsgBox "�L�����Z�����܂���", vbCritical

End If

Set Cnn = Nothing
Set Cmd = Nothing
Set Rs = Nothing

Application.StatusBar = Old_Status


End Sub
