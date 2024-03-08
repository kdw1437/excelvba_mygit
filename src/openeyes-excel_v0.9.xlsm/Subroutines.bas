Attribute VB_Name = "Subroutines"
' 주어진 data string과 함께 POST request를 특정 URL에 보낸다. response는 message box에 보여진다.
'
' @subroutine SendPostRequest
' @param {String} DataString - POST request에서 보내질 데이터
' @param {String} url - POST request가 보내질 URL
Sub SendPostRequest(DataString As String, url As String)
    Dim xmlhttp As Object
    
    ' 새로운 XML HTTP request를 만든다.
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' xmlhttp 객체를 세팅한다. 특정 URL에 synchronous하게 POST request를 한다. xmlhttp객체의 내부 상태를 configure한다.
    xmlhttp.Open "POST", url, False
    
    ' request content-type header를 application/x-www-form-urlencoded로 세팅한다.
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

'Internal State는 모든 field를 포함한다. 하지만, 직접적으로 개개의 field에 mapping되지 않는 부가적인 상태 정보도 포함한다. (예) 메소드로 관리되어지는 flag, counter, state variable
'객체의 Internal state는 현재 데이터(field)와 data가 메서드에 의해서 조작되어진 역사 및 맥락의 결합이다.
