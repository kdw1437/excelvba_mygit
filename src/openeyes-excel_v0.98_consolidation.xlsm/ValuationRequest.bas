Attribute VB_Name = "ValuationRequest"
Sub ValuationRequest()
    Dim jsonString As String
    Dim xmlhttp As Object
    Dim baseurl As String
    
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
    baseurl = "http://urosys-web.juroinstruments.com/app/"
    url = "http://urosys-web.juroinstruments.com/app/createValWebJob"
    url = baseurl & "createValWebJob"
    
    xmlhttp.Open "POST", url, False
    
    xmlhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    xmlhttp.Send sendMessage
    
    Do
        
        If xmlhttp.Status = 200 Then
        
            Exit Do
        
        ElseIf xmlhttp.Status >= 400 Then
            
            'ThisWorkbook.Sheets("Sheet1").Range("F2").value = xmlhttp.Status
            ws.Range("C5").value = xmlhttp.Status
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
    
    'ThisWorkbook.Sheets("Sheet1").Range("B5").value = jobId
    ws.Range("B5").value = jobId
    
    Dim url2 As String
    url2 = baseurl & "selectValJob?jobId=" & jobId
    
    Dim startTime As Single
    Dim elapsedTime As Single
    
    startTime = Timer
    elapsedTime = 0
    
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
                        
            ws.Range("E5").value = "'" & ParsedJson("procEndDtime")
            
            ws.Range("C5").value = ParsedJson("jobStateCodeNm")
            Exit Do
            
        ElseIf ParsedJson("jobStateCode") = "F" Or ParsedJson("jobStateCode") = "C" Then
        
            ws.Range("C5").value = ParsedJson("jobStateCodeNm")
            
            ws.Range("E5").value = "'" & ParsedJson("procEndDtime")
            Exit Do
            'Exit Sub
        End If
        
        ws.Range("C5").value = ParsedJson("jobStateCodeNm")
        ws.Range("D5").value = "'" & ParsedJson("creDtime")
        
        'Start timing (StartTime은 Timer의 현재 시간을 assign받고,
        'startTime = Timer
        
        elapsedTime = Timer - startTime
        
        If elapsedTime >= 10 Then
            Exit Do
        End If
        
        Dim loopStartTime As Single
        loopStartTime = Timer
        Do While Timer < loopStartTime + 10
            DoEvents
        Loop
        '10초 마다 loop
'        Do While Timer < startTime + 10
'            DoEvents  'Excel이 responsive하도록
'        Loop
        DoEvents
    Loop
    
    If ParsedJson("jobStateCode") = "FIN" Then
        Dim url3 As String
        url3 = baseurl & "SelectJob1?jobid=" & jobId
        
        Dim xmlhttp3 As Object
        Set xmlhttp3 = CreateObject("WinHttp.WinHttpRequest.5.1")
        
        xmlhttp3.Open "GET", url3, False
        xmlhttp3.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        xmlhttp3.Send
        
        response = xmlhttp3.ResponseText
        Dim ParsedJson2 As Dictionary
        Set ParsedJson2 = JsonConverter.ParseJson(response)
        
        Dim selectjob1 As Variant
        
        Set selectjob1 = ParsedJson2("selectjob1")
        
        Dim rowIndex As Integer
    
        rowIndex = 10
    
        Dim job As Variant
        Dim currentRow As Integer
        Dim itemCd As String
                
        Do While ws.Cells(rowIndex, 2).value <> ""
            itemCd = ws.Cells(rowIndex, 2).value
            
            For Each job In selectjob1
                If job("itemCd") = itemCd Then
                    ws.Cells(rowIndex, 32).value = "FIN"
                    ws.Cells(rowIndex, 33).value = job("price")
                    Exit For
                End If
                DoEvents
            Next job
    
            rowIndex = rowIndex + 1
            DoEvents
        Loop
    
    Else
        
        Dim url4 As String
        url4 = baseurl & "selectJobFail?jobId=" & jobId
    
        Dim xmlhttp4 As Object
        Set xmlhttp4 = CreateObject("WinHttp.WinHttpRequest.5.1")
        
        xmlhttp4.Open "GET", url4, False
        xmlhttp4.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        xmlhttp4.Send
        
        response = xmlhttp4.ResponseText
        Dim ParsedJson4 As Dictionary
        Set ParsedJson4 = JsonConverter.ParseJson(response)
        
        Dim selectjob4 As Variant
        
        Set selectjob4 = ParsedJson4("selectJobFail")
        
        Dim rowIndex4 As Integer
        rowIndex4 = 10
        
        For Each job In selectjob4
            
            ws.Cells(rowIndex4, 32).value = job("taskSttsCd")
            'ws.Cells(rowIndex4, 33).value = job("price")
            rowIndex4 = rowIndex4 + 1
            
            DoEvents
        Next job
    
    End If
        
End Sub

Sub FindColumnNumber()

    Dim columnLetter As String
    Dim columnNumber As Integer
    
    columnLetter = "AG"
    
    columnNumber = Range(columnLetter & "1").Column
    
    Debug.Print columnNumber

End Sub
