Attribute VB_Name = "Module1"
'���t�[�V���b�s���OWebAPI���g���āAJAN�Ō������ĉ��i�ƃV���b�v���𔲂��o��
'GET���\�b�h�œ���JAN�̉��i���ȂǁA�������ʂ�XML���擾�ł���̂ŁA������p�[�X���ăZ���ɓ]�L����

'�Q�Ɛݒ�ŁAMicroSoft XML,V6.0�@���C�u�����Ƀ`�F�b�N�����邱��
'�G�N�Z����MXL���������߂̃��C�u�����AMSXML2�I�u�W�F�N�g�̐����ɕK�v

'���t�[�V���b�s���OWebAPI���t�@�����X�@WebAPI���Ăяo���ɂ͗v�A�v���P�[�V�����R�[�h
'http://developer.yahoo.co.jp/webapi/shopping/shopping/v1/itemsearch.html

'MSDN MSXML2.XMLHTTP�I�u�W�F�N�g�̌������t�@�����X�͉��L
'http://msdn.microsoft.com/en-us/library/ms759148%28v=vs.85%29.aspx

'MSDN ���S�҂̂��߂� XML DOM �K�C�h�@�X�V�������Â����Ǌ�{�͕ς��Ȃ��͂�
'http://msdn.microsoft.com/ja-jp/library/aa468547.aspx


'�錾�Z�N�V����

'sleep���g�����߂̐錾
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'�S�v���V�[�W���ŋ��L����ϐ�
'�V���b�v/���i���X�g���L������J�n��
Dim startcolumn_price As Integer

'Yahoo API�Ăяo���Ɏg�� �A�v���P�[�V����ID
Const APP_ID = ""

'jan�̓����Ă���Z��
Dim c As Range

Sub ���i����()

'�X�ܖ��E���i�������o���ŏ��̗���w��
startcolumn_price = ActiveSheet.UsedRange.Columns.Count + 1

'Jan�̗����肵�܂��B
'�Г��V�X�e���ł́uJAN�R�[�h�vYahooCsv�ł́ujan�v�Ȃ̂łǂ��炩�𒲂ׂ�
Dim jan_col As Range

If Not Range("1:1").Find("jan") Is Nothing Then
    
    Set jan_col = Range("1:1").Find("jan")

ElseIf Not Range("1:1").Find("JAN�R�[�h") Is Nothing Then

    Set jan_col = Range("1:1").Find("JAN�R�[�h")

Else
    
    MsgBox "Jan�̌��o����������܂���B" & vbLf & _
           "1�s�ڂ̃w�b�_�[�Ɂujan�v���uJAN�R�[�h�v���w��"
    
    Exit Sub

End If

'JAN�̃��X�g�������W�Ŏ擾���܂�

Dim rng_jan As Range

'rng_jan��JAN�R�[�h�̃����W���Z�b�g
Set rng_jan = jan_col.Offset(1, 0).Resize(ActiveSheet.UsedRange.Rows.Count - 1, 1)

For Each c In rng_jan
       
    '�T�[�o�[�A�N�Z�X�Ƃ��A���W���[����or�I�u�W�F�N�g�����ĕ��񏈗��ł���Ƒ�����ł́H
    '���̃R�[�h�ł��\���������ǁc���t�[�T�[�o�[�̃��X�|���X�������A�����Ă�WEB�|�[�^��
    
    Dim jan As String
    jan = c.Value
    
    '�Z������擾����jan������13�P�^���`�F�b�N
    If Not jan Like "#############" Then
        Call writeError("�������Ȃ�JAN")
        GoTo continue:
    End If

    Call loadXml(jan)
    
continue:

Next c

End Sub

Private Sub loadXml(jan As String)

 'xml�I�u�W�F�N�g�̃C���X�^���X����
 Dim xml As Object
 Set xml = CreateObject("MSXML2.DOMDocument")
 
 '�T�[�o�[����XML��ǂނ��߂�MSXML2�I�u�W�F�N�g�̐ݒ�
 xml.async = False
 xml.setProperty "ServerHTTPRequest", True
 
 'GET���邽�߂�url�𐶐����܂�
 Dim url As String
 url = makeUrl(jan)
 
 '�T�[�o�[��GET���\�b�h�ŃA�h���X�𓊂��āAXML���擾���܂�
 xml.Load (url)

 '�X���[�v�^�C���̃J�E���^j��������
 j = 0
 
 '�T�[�o�[����̃��X�|���X�҂��@Sleep100���Ȃ���ҋ@�@�r�W�[�E�F�C�g
 Do
     DoEvents
     Sleep 10
     j = j + 1
     
     If j > 100 Then  '10msec*100=1�b�������Ȃ���΃��[�v�A�E�g
         Cells(c.Row, startcolumn_price).Value = "�T�[�o�[�����Ȃ�"
         Exit Do
     End If
         
 Loop While xml.readyState <> 4
         
 'xml�I�u�W�F�N�g��XML���擾�ł��ĂȂ���΁AContinue
 
 If Not xml.HasChildNodes Then
     Call writeError("���ʂ��������擾�ł��܂���ł���")
     Exit Sub
 End If

    
'xml��ResultSet����Hit�����o���A�ȒP�ȃc���[�\��
'ResultSet>Result>Hit>Store>Name
'                    >Price

'jan�������w��Ȃ�����<Error><Message>BadRequest���Ԃ��Ă���̂ŁAResultset�����邩�`�F�b�N
If xml.getElementsByTagName("ResultSet").Length > 0 Then
    
    'ResultSet��TotalResultAvailable/TotalResult����
    '����ł͓��ɃZ���ɏ����߂��Ȃ����A�L����Hit�v�f���̃`�F�b�N�ƁA
    '�����\�ȓX�ܐ��E�f�ړX�ܐ����c���ł���̂ŕϐ��Ɋi�[����
    'TotalResultsReturned=0���Ƌ��Hit�v�f��1�Ԃ��Ă���
    
    Dim total_results_counts As Integer
    total_results_counts = xml.SelectSingleNode("ResultSet").Attributes.getNamedItem("totalResultsReturned").Text
    
    Dim available_results_counts As Integer
    available_results_counts = xml.SelectSingleNode("ResultSet").Attributes.getNamedItem("totalResultsAvailable").Text
    
    If total_results_counts > 0 Then
        
        Call parseWriteRanking(xml.SelectNodes("ResultSet/Result/Hit"))
    
    Else
        'totalRsultsAvailable��0
        Call writeError("�f�ڃV���b�v�Ȃ�")
        Exit Sub
    
    End If
    
Else

    'Resultset���Ȃ�
    Call writeError("�Y��JAN�����t�[�ɓo�^�Ȃ�")
    Exit Sub
    
End If

End Sub

Private Function makeUrl(jan As String)
'jan��n���Ă�����āAGET���\�b�h�œ�����URL�𐶐�

Dim base_url As String 'WEB API��GET�ŌĂяo���x�[�XURL
base_url = "http://shopping.yahooapis.jp/ShoppingWebService/V1/itemSearch"

Dim sort As String
sort = "%2Bprice" '���i���A�{�|�ō~���E�����w��ł���AURL�G���e�B�e�B�ɕϊ����K�v

get_url = base_url
get_url = get_url & "?appid=" & APP_ID '�A�v���P�[�V����ID
get_url = get_url & "&sort=" & sort    '�\�[�g���
get_url = get_url & "&hits=5"          '�ő吔5�A�ň�5����
get_url = get_url & "&jan=" & jan      'JAN���Z�b�g
    
makeUrl = get_url

End Function

Private Sub parseWriteRanking(hits As Object)
'<Hit>�m�[�h���X�ghits���C�e���[�g���Z���ɏ�������

'h�����������Ȃ���C�e���[�g����̂ŁAhit�m�[�h���i�[�ł���ϐ�h���Z�b�g�A
'h��XML DOM Element��Node�N���X�ANodeList�͕����m�[�h���i�[����N���X
Dim h As Object

'�����߂��Z���̗�J�E���^�[k��ݒ�
Dim k As Integer
k = 0

For Each h In hits

    store_name = h.SelectSingleNode("Store/Name").Text              '�eHit�m�[�h��Store>Name���V���b�v��
    
    Cells(c.Row, startcolumn_price + k).Value = store_name          '���+1���Ȃ���V���b�v�E���i�E�V���b�v�E���i�̏���Cell�ɋL��
    
    k = k + 1
            
    sale_price = h.SelectSingleNode("Price").Text                   '�eHit�m�[�h��Price���̔����i
    Cells(c.Row, startcolumn_price + k).Value = sale_price
    k = k + 1
    
Next h

End Sub

Private Sub writeError(s As String)
'�G���[���b�Z�[�W���Z���ɕԂ�

Cells(c.Row, startcolumn_price).Value = s

End Sub
