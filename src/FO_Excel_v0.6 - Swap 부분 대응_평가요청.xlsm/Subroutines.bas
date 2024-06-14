Attribute VB_Name = "Subroutines"
' �־��� data string�� �Բ� POST request�� Ư�� URL�� ������. response�� message box�� ��������.
'
' @subroutine SendPostRequest
' @param {String} DataString - POST request���� ������ ������
' @param {String} url - POST request�� ������ URL
Sub SendPostRequest(DataString As String, url As String)
    Dim xmlhttp As Object
    
    ' ���ο� XML HTTP request�� �����.
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' xmlhttp ��ü�� �����Ѵ�. Ư�� URL�� synchronous�ϰ� POST request�� �Ѵ�. xmlhttp��ü�� ���� ���¸� configure�Ѵ�.
    xmlhttp.Open "POST", url, True
    
    ' request content-type header�� application/x-www-form-urlencoded�� �����Ѵ�.
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    ' DataString�� �Բ� request�� ������.
    xmlhttp.Send "a=" & DataString
    
    ' request�� �Ϸ�� ������ ����Ѵ�.
    Do While xmlhttp.readyState <> 4
        DoEvents
    Loop
    
    ' request�� ���¸� Ȯ���Ѵ�.
    If xmlhttp.Status = 200 Then
        ' request�� �������̾��ٸ�, response�� ����Ѵ�.
        MsgBox xmlhttp.responseText
    Else
        ' request�� �����ߴٸ�, ���¸� ����Ѵ�.
        MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.statusText
    End If
    
    ' ����
    Set xmlhttp = Nothing
End Sub

'Internal State�� ��� field�� �����Ѵ�. ������, ���������� ������ field�� mapping���� �ʴ� �ΰ����� ���� ������ �����Ѵ�. (��) �޼ҵ�� �����Ǿ����� flag, counter, state variable
'��ü�� Internal state�� ���� ������(field)�� data�� �޼��忡 ���ؼ� ���۵Ǿ��� ���� �� �ƶ��� �����̴�.




