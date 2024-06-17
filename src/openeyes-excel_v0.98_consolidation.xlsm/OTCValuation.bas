Attribute VB_Name = "OTCValuation"
Option Explicit

Sub OTCValuation()
    Dim JsonString As String
    Dim xmlhttp As Object
    
    ' code에서 JSON string을 바로 정의한다.
    'JsonString = StrConv("test=test1", vbFromUnicode)
    
    ' 새로운 XML HTTP request를 만든다.
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' 평가를 요청하는 url
    Dim url As String
    url = "http://urosys-web.juroinstruments.com/app/createValWebJob"
    
    Dim baseurl As String
    baseurl = "http://urosys-web.juroinstruments.com/app/"
    
    ' post를 위한 xmlhttp의 메소드 Open 호출
    xmlhttp.Open "POST", url, False
    
    xmlhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("OTC")
    
    Dim valDate As String
    valDate = Format(ws.Range("B2").value, "yyyymmdd")
    
    Dim name As String
    name = ws.Range("A2").value
    
    Dim valTypeCode As String
    valTypeCode = ws.Range("C2").value
    
    Dim contextIds As String
    contextIds = ws.Range("D2").value
    
    Dim officeCd As String
    officeCd = ws.Range("E2").value
    
    Dim priority As String
    priority = CStr(ws.Range("F2").value)
    
    Dim itemCodes As String
    itemCodes = ""
    
    Dim lastRow As Integer
    lastRow = ws.Cells(9, 1).End(xlDown).row
    
    Dim i As Integer
    
    For i = 9 To lastRow
        itemCodes = itemCodes & ws.Cells(i, 1).value & ","
    Next i
    
    ' 마지막 콤마 제거
    If Len(itemCodes) > 0 Then
        itemCodes = Left(itemCodes, Len(itemCodes) - 1)
    End If
    
    ' request를 보낸다.
    'xmlhttp.Send "officeCd=BO&name=TEST4&valDate=20231228&valTypeCode=P&greekLevel=&contextIds=BO&dataSetIds=Test_4,official&simId=&priority=4&itemCodes=ELS3588"
    Dim sendMessage As String
    sendMessage = "officeCd=" + officeCd + "&name=" + name + "&valDate=" + valDate + "&valTypeCode=" + valTypeCode + "&greekLevel=&contextIds=" + contextIds + "&dataSetIds=official&simId=&priority=" + priority + "&itemCodes=" + itemCodes
    
    Debug.Print sendMessage
    
    xmlhttp.Send sendMessage
    
    Do
        
        If xmlhttp.Status = 200 Then
        
            Exit Do
        
        ElseIf xmlhttp.Status >= 400 Then
            
            ws.Range("K2").value = xmlhttp.Status
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
    
    ws.Range("B5").value = jobId
    
    Dim url2 As String
    url2 = baseurl & "selectValJob?jobId=" & jobId
    
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
            ws.Range("C5").value = ParsedJson("jobStateCode")
            
            ws.Range("E5").value = "'" & ParsedJson("procEndDtime")
            Exit Do
            
        ElseIf ParsedJson("jobStateCode") = "F" Or ParsedJson("jobStateCode") = "C" Then
        
            ws.Range("C5").value = ParsedJson("jobStateCode")
            
            ws.Range("E5").value = "'" & ParsedJson("procEndDtime")
            Exit Sub
            
        End If
        
        ws.Range("C5").value = ParsedJson("jobStateCode")
        ws.Range("D5").value = "'" & ParsedJson("creDtime")
        
        'Start timing (StartTime은 Timer의 현재 시간을 assign받고,
        Dim startTime As Single
        startTime = Timer
        
        '10초 마다 loop
        Do While Timer < startTime + 10
            DoEvents  'Excel이 responsive하도록
        Loop
        
        If ParsedJson("jobStateCode") = "W" Then
            ws.Range("C5").value = ParsedJson("jobStateCode")
            Exit Sub
        End If
        DoEvents
    Loop
    
    Dim url3 As String
    url3 = baseurl & "SelectJob1?jobid=" & jobId
    
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
    
    rowIndex = 9
    
    Dim Item As Dictionary
    
    For Each Item In selectjob1
        Dim foundCell As Range
        Set foundCell = ws.Columns("A").Find(What:=Item("itemCd"), LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundCell Is Nothing Then
            foundCell.Offset(0, 1).value = Item("price")
        End If
        DoEvents
    Next Item
    
    ImportGreekData jobId
    
End Sub

               
'평가작업'

