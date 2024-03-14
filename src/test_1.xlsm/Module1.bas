Attribute VB_Name = "Module1"
Sub SendJSON()
    Dim JsonString As String
    Dim xmlhttp As Object
    
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    Dim url As String
    url = "http://urosys-web.juroinstruments.com/app/createValWebJob"
    
    xmlhttp.Open "POST", url, False
    
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    xmlhttp.Send "officeCd=BO&name=TEST4&valDate=20231228&valTypeCode=P&greekLevel=&contextIds=BO&dataSetIds=Test_4,official&simId=&priority=4&itemCodes=ELS3588"
    
    Do
        
        If xmlhttp.Status = 200 Then
        
            Exit Do
        
        ElseIf xmlhttp.Status >= 400 Then
            
            Exit Sub
        
        End If
        
    Loop
    
    Dim response As String
    response = xmlhttp.ResponseText
    
    Dim jobId As String
    Dim ParsedJson As Dictionary
    
    Set ParsedJson = JsonConverter.ParseJson(response)
    jobId = ParsedJson("jobId")
    
    Dim url2 As String
    url2 = "http://urosys-web.juroinstruments.com/app/selectValJob?jobId=" & jobId
    
    Do
        Dim xmlhttp2 As Object
        Set xmlhttp2 = CreateObject("WinHttp.WinHttpRequest.5.1")
        
        xmlhttp2.Open "GET", url2, False
        xmlhttp2.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        xmlhttp2.Send
        
        response = xmlhttp2.ResponseText
        Set ParsedJson = JsonConverter.ParseJson(response)
        
        If ParsedJson("jobStateCodeNm") = "¿Ï·á" Then
            ' If the job is complete, continue with the logic
            Exit Do
        End If
        
        Application.Wait Now + TimeValue("00:00:10")
    Loop
    
End Sub
