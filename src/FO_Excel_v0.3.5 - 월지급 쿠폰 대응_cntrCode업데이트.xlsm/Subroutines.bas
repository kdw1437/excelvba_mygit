Attribute VB_Name = "Subroutines"
' 주어진 data string과 함께 POST request를 특정 URL에 보낸다. response는 message box에 보여진다.
'
' @subroutine SendPostRequest
' @param {String} DataString - POST request에서 보내질 데이터
' @param {String} url - POST request가 보내질 URL
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


'Internal State는 모든 field를 포함한다. 하지만, 직접적으로 개개의 field에 mapping되지 않는 부가적인 상태 정보도 포함한다. (예) 메소드로 관리되어지는 flag, counter, state variable
'객체의 Internal state는 현재 데이터(field)와 data가 메서드에 의해서 조작되어진 역사 및 맥락의 결합이다.



