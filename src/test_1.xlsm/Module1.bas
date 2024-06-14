Attribute VB_Name = "Module1"
Sub SendJSON()
    Dim JsonString As String
    Dim xmlhttp As Object
    Dim baseUrl As String
    
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    baseUrl = ThisWorkbook.Sheets("Sheet1").Range("A8").Value
    
    Dim url As String
    'url = "http://urosys-web.juroinstruments.com/app/createValWebJob"
    url = baseUrl & "createValWebJob"
    
    xmlhttp.Open "POST", url, False
    
    xmlhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    xmlhttp.Send "officeCd=BO&name=TEST4&valDate=20231228&valTypeCode=P&greekLevel=&contextIds=BO&dataSetIds=Test_4,official&simId=&priority=4&itemCodes=ELS3588"
    
    Do
        
        If xmlhttp.Status = 200 Then
        
            Exit Do
        
        ElseIf xmlhttp.Status >= 400 Then
            
            ThisWorkbook.Sheets("Sheet1").Range("F2").Value = xmlhttp.Status
            Exit Sub
        
        End If
        
        DoEvents
    Loop
    
    Dim response As String
    response = xmlhttp.ResponseText
    
    'status 완료 확인 (status가 완료 시에, ResponseText로 부터 jobId를 획득해 올 수 있다.)
    'jobId 획득
    
    Dim jobId As String
    Dim ParsedJson As Dictionary
    
    Set ParsedJson = JsonConverter.ParseJson(response)
    jobId = ParsedJson("jobId")
    
    ThisWorkbook.Sheets("Sheet1").Range("B5").Value = jobId
    
    Dim url2 As String
    url2 = baseUrl & "selectValJob?jobId=" & jobId
    
    Do
        Dim xmlhttp2 As Object
        Set xmlhttp2 = CreateObject("WinHttp.WinHttpRequest.5.1")
        
        xmlhttp2.Open "GET", url2, False
        xmlhttp2.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        xmlhttp2.Send
        
        response = xmlhttp2.ResponseText
        
        Set ParsedJson = JsonConverter.ParseJson(response)
        
        If ParsedJson("jobStateCode") = "FIN" Then
            ' job이 완료되면, logic을 지속함.
            ThisWorkbook.Sheets("Sheet1").Range("C5").Value = ParsedJson("jobStateCode")
            
            ThisWorkbook.Sheets("Sheet1").Range("E5").Value = "'" & ParsedJson("procEndDtime")
            Exit Do
            
        ElseIf ParsedJson("jobStateCode") = "F" Or ParsedJson("jobStateCode") = "C" Then
        
            ThisWorkbook.Sheets("Sheet1").Range("C5").Value = ParsedJson("jobStateCode")
            
            ThisWorkbook.Sheets("Sheet1").Range("E5").Value = "'" & ParsedJson("procEndDtime")
            Exit Sub
            
        End If
        
        ThisWorkbook.Sheets("Sheet1").Range("C5").Value = ParsedJson("jobStateCode")
        ThisWorkbook.Sheets("Sheet1").Range("D5").Value = "'" & ParsedJson("creDtime")
        
        'Start timing (StartTime은 Timer의 현재 시간을 assign받고,
        startTime = Timer
        
        '10초 마다 loop
        Do While Timer < startTime + 10
            DoEvents  'Excel이 responsive하도록
        Loop
    Loop
    
    Dim url3 As String
    url3 = baseUrl & "SelectJob1?jobid=" & jobId
    
    Dim xmlhttp3 As Object
    Set xmlhttp3 = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    xmlhttp3.Open "GET", url3, False
    xmlhttp3.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xmlhttp3.Send
    
    response = xmlhttp3.ResponseText
    
    Set ParsedJson = JsonConverter.ParseJson(response)
    
    Dim selectjob1 As Variant
    
    Set selectjob1 = ParsedJson("selectjob1")
    
    Dim rowIndex As Integer
    
    rowIndex = 1
    
    Dim job As Dictionary
    
    For Each job In selectjob1
        ThisWorkbook.Sheets("Sheet1").Cells(rowIndex, 10).Value = job("jobId")
        ThisWorkbook.Sheets("Sheet1").Cells(rowIndex, 11).Value = job("price")
        rowIndex = rowIndex + 1
        
        DoEvents
    Next job
End Sub
