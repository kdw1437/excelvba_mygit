Attribute VB_Name = "OTCValuation"
Option Explicit

Sub OTCValuation()
    Dim JsonString As String
    Dim xmlhttp As Object
    
    ' code���� JSON string�� �ٷ� �����Ѵ�.
    'JsonString = StrConv("test=test1", vbFromUnicode)
    
    ' ���ο� XML HTTP request�� �����.
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' �򰡸� ��û�ϴ� url
    Dim url As String
    url = "http://urosys-web.juroinstruments.com/app/createValWebJob"
    
    Dim baseurl As String
    baseurl = "http://urosys-web.juroinstruments.com/app/"
    
    ' post�� ���� xmlhttp�� �޼ҵ� Open ȣ��
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
    
    ' ������ �޸� ����
    If Len(itemCodes) > 0 Then
        itemCodes = Left(itemCodes, Len(itemCodes) - 1)
    End If
    
    ' request�� ������.
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
    
    'status �Ϸ� Ȯ�� (status�� �Ϸ� �ÿ�, ResponseText�� ���� jobId�� ȹ���� �� �� �ִ�.)
    'jobId ȹ��
    
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
            ' job�� �Ϸ�Ǹ�, logic�� ������.
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
        
        'Start timing (StartTime�� Timer�� ���� �ð��� assign�ް�,
        Dim startTime As Single
        startTime = Timer
        
        '10�� ���� loop
        Do While Timer < startTime + 10
            DoEvents  'Excel�� responsive�ϵ���
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

               
'���۾�'

