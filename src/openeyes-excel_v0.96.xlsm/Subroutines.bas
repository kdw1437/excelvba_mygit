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
    xmlhttp.Open "POST", url, False
    
    ' request content-type header�� application/x-www-form-urlencoded�� �����Ѵ�.
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    ' Send the request with the DataString
    xmlhttp.Send "a=" & DataString
    
    ' Check the status of the request
    If xmlhttp.Status = 200 Then
        ' If the request was successful, output the response
        MsgBox xmlhttp.responseText
    Else
        ' If the request failed, output the status
        MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.statusText
    End If
    
    ' Clean up
    Set xmlhttp = Nothing
End Sub

'Internal State�� ��� field�� �����Ѵ�. ������, ���������� ������ field�� mapping���� �ʴ� �ΰ����� ���� ������ �����Ѵ�. (��) �޼ҵ�� �����Ǿ����� flag, counter, state variable
'��ü�� Internal state�� ���� ������(field)�� data�� �޼��忡 ���ؼ� ���۵Ǿ��� ���� �� �ƶ��� �����̴�.


' �־��� data string�� �Բ� POST request�� Ư�� URL�� ������. response�� message box�� ��������.
'
' @subroutine SendPostRequestAsync
' @param {String} DataString - POST request���� ������ ������
' @param {String} url - POST request�� ������ URL
Sub SendPostRequestAsync(DataString As String, url As String)
    Dim xmlhttp As Object
    
    ' ���ο� XML HTTP request�� �����.
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' xmlhttp ��ü�� �����Ѵ�. Ư�� URL�� asynchronous�ϰ� POST request�� �Ѵ�.
    xmlhttp.Open "POST", url, True
    
    ' request content-type header�� application/x-www-form-urlencoded�� �����Ѵ�.
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    ' Event handlers to handle the response and state changes
    ' Note: These handlers need to be added to your VBA project in a class module.
    
    ' Send the request with the DataString
    xmlhttp.Send "a=" & DataString
    
    ' Wait for the response
    Do While xmlhttp.readyState <> 4
        DoEvents ' Keep Excel responsive
    Loop
    
    ' Check the status of the request
    If xmlhttp.Status = 200 Then
        ' If the request was successful, output the response
        MsgBox xmlhttp.responseText
    Else
        ' If the request failed, output the status
        MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.statusText
    End If
    
    ' Clean up
    Set xmlhttp = Nothing
End Sub

