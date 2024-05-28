Attribute VB_Name = "Subroutines"
' �־��� data string�� �Բ� POST request�� Ư�� URL�� ������. response�� message box�� ��������.
'
' @subroutine SendPostRequest
' @param {String} DataString - POST request���� ������ ������
' @param {String} url - POST request�� ������ URL
Sub SendPostRequest(DataString As String, url As String)
    Dim xmlhttp As Object
    Dim responseJson As Object
    Dim i As Long

    ' Create a new XML HTTP request
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Configure the xmlhttp object
    xmlhttp.Open "POST", url, True
    
    ' Set the request content-type header to application/x-www-form-urlencoded
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    ' Send the request with the DataString
    xmlhttp.Send "a=" & DataString
    
    ' Check the status of the request
    If xmlhttp.Status = 200 Then
        ' If the request was successful, parse the JSON response
        Set responseJson = JsonConverter.ParseJson(xmlhttp.responseText)
        'Debug.Print responseJson
        ' Assuming responseJson is an array of objects
        If IsObject(responseJson) And TypeName(responseJson) = "Collection" Then
            For i = 1 To responseJson.Count
                ' Print each cntrCode to the Immediate Window
                Debug.Print "cntrCode: " & responseJson(i)("cntrCode")
            Next i
        Else
            MsgBox "Unexpected JSON format in response"
        End If
    Else
        ' If the request failed, output the status
        MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.statusText
    End If
    
    ' Clean up
    Set xmlhttp = Nothing
End Sub


'Internal State�� ��� field�� �����Ѵ�. ������, ���������� ������ field�� mapping���� �ʴ� �ΰ����� ���� ������ �����Ѵ�. (��) �޼ҵ�� �����Ǿ����� flag, counter, state variable
'��ü�� Internal state�� ���� ������(field)�� data�� �޼��忡 ���ؼ� ���۵Ǿ��� ���� �� �ƶ��� �����̴�.



