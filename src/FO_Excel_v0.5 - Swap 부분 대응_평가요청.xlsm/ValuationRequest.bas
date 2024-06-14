Attribute VB_Name = "ValuationRequest"
Sub ValuationRequest()
    Dim JsonString As String
    Dim xmlhttp As Object
    Dim baseUrl As String
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Quote")
    
    Dim requestName As String
    requestName = "Quote Valuation By " + ws.Cells(2, 5).value
    
    Dim valDate As String
    valDate = Format(ws.Cells(2, 1).value, "yyyymmdd")
    
    ',(comma)를 이용해서 itemCode(cntrCode)를 가져오도록 하기
    Dim itemCodes As String
    itemCodes = ""
    
    Dim lastRow As Integer
    lastRow = ws.Cells(10, 2).End(xlDown).row
    
    For i = 10 To lastRow
        itemCodes = itemCodes & ws.Cells(i, 2).value & ","
    Next i
    
    ' 마지막 콤마 제거
    If Len(itemCodes) > 0 Then
        itemCodes = Left(itemCodes, Len(itemCodes) - 1)
    End If
    
    Dim sendMessage As String
    sendMessage = "officeCd=FO&name=" + requestName + "&valDate=" + valDate + "&valTypeCode=P&greekLevel=&contextIds=FO&dataSetIds=official&simId=&priority=4&itemCodes=" + itemCodes
    
    Debug.Print sendMessage
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
'    baseUrl = ThisWorkbook.Sheets("Sheet1").Range("A8").value
    
    Dim url As String
    baseUrl = "http://urosys-web.juroinstruments.com/app/"
    url = "http://urosys-web.juroinstruments.com/app/createValWebJob"
    url = baseUrl & "createValWebJob"
    
    xmlhttp.Open "POST", url, False
    
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets("Quote")
'
'    Dim requestName As String
'    requestName = "Quote Valuation By " + Worksheets("Quote").Cells(2, 5).value
    
'    Dim sendMessage As String
'    sendMessage = "officeCd=FO&name=" + requestName + "&valDate=20231228&valTypeCode=P&greekLevel=&contextIds=BO&dataSetIds=Test_4,official&simId=&priority=4&itemCodes=ELS3588"
    
    xmlhttp.Send sendMessage
    
    Do
        
        If xmlhttp.Status = 200 Then
        
            Exit Do
        
        ElseIf xmlhttp.Status >= 400 Then
            
            'ThisWorkbook.Sheets("Sheet1").Range("F2").value = xmlhttp.Status
            'ws.Range("F2").value = xmlhttp.Status
            Exit Sub
        
        End If
        
        DoEvents
    Loop
    
    Dim response As String
    response = xmlhttp.responseText
    
    'status 완료 확인 (status가 완료 시에, ResponseText로 부터 jobId를 획득해 올 수 있다.)
    'jobId 획득
    
    Dim jobId As String
    Dim ParsedJson As Dictionary
    
    Set ParsedJson = JsonConverter.ParseJson(response)
    jobId = ParsedJson("jobId")
    
    'ThisWorkbook.Sheets("Sheet1").Range("B5").value = jobId
    ws.Range("B5").value = jobId
    
    Dim url2 As String
    url2 = baseUrl & "selectValJob?jobId=" & jobId
    
    Do
        Dim xmlhttp2 As Object
        Set xmlhttp2 = CreateObject("WinHttp.WinHttpRequest.5.1")
        
        xmlhttp2.Open "GET", url2, False
        xmlhttp2.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        xmlhttp2.Send
        
        response = xmlhttp2.responseText
        
        Set ParsedJson = JsonConverter.ParseJson(response)
        
        If ParsedJson("jobStateCode") = "FIN" Then
            ' job이 완료되면, logic을 지속함.
            'ThisWorkbook.Sheets("Sheet1").Range("C5").value = ParsedJson("jobStateCode")
            
            ''ThisWorkbook.Sheets("Sheet1").Range("D5").value = "'" & ParsedJson("procEndDtime")
            Exit Do
            
        ElseIf ParsedJson("jobStateCode") = "F" Or ParsedJson("jobStateCode") = "C" Then
        
            'ThisWorkbook.Sheets("Sheet1").Range("C5").value = ParsedJson("jobStateCode")
            
            ''ThisWorkbook.Sheets("Sheet1").Range("D5").value = "'" & ParsedJson("procEndDtime")
            Exit Sub
            
        End If
        
        'ThisWorkbook.Sheets("Sheet1").Range("C5").value = ParsedJson("jobStateCode")
        ''ThisWorkbook.Sheets("Sheet1").Range("C5").value = "'" & ParsedJson("creDtime")
        
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
    xmlhttp3.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xmlhttp3.Send
    
    response = xmlhttp3.responseText
    
    Set ParsedJson = JsonConverter.ParseJson(response)
    
    Dim selectjob1 As Variant
    
    Set selectjob1 = ParsedJson("selectjob1")
    
'    Dim rowIndex As Integer
'
'    rowIndex = 1
'
'    Dim job As Dictionary
'
'    For Each job In selectjob1
'        ThisWorkbook.Sheets("Sheet1").Cells(rowIndex, 10).value = job("jobId")
'        ThisWorkbook.Sheets("Sheet1").Cells(rowIndex, 11).value = job("price")
'        rowIndex = rowIndex + 1
'
'        DoEvents
'    Next job
End Sub

