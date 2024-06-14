Attribute VB_Name = "AppWait"
'�� ������� �� ��, ������ �߻��Ѵ�. excel�� responsive�� ���� ����ٰ� ���ƿͼ� ���� �߻�. �׳� Forloop�� doevent�� ����ϴ� Module1�� ����ϴ� ���� ����.
Sub SendJSON()
    Dim JsonString As String
    Dim xmlhttp As Object
    
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    Dim url As String
    url = "http://urosys-web.juroinstruments.com/app/createValWebJob"
    
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
    
    'status �Ϸ� Ȯ�� (status�� �Ϸ� �ÿ�, ResponseText�� ���� jobId�� ȹ���� �� �� �ִ�.)
    'jobId ȹ��
    
    Dim jobId As String
    Dim ParsedJson As Dictionary
    
    Set ParsedJson = JsonConverter.ParseJson(response)
    jobId = ParsedJson("jobId")
    
    ThisWorkbook.Sheets("Sheet1").Range("B5").Value = jobId
    
    Dim url2 As String
    url2 = "http://urosys-web.juroinstruments.com/app/selectValJob?jobId=" & jobId
    
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
        
        Dim endTime As Date
        endTime = Now + TimeValue("00:00:10") ' ������ �ð��� �������� ���� 10�� �ķ� �����Ѵ�.
        
        Do While Now < endTime
            DoEvents ' ������ responsive�ϵ��� �����Ѵ�.
            ' CPU ����� ���̱� ���� �ſ� ª�� �ð� ��ٸ�: ���û���
            Application.Wait (Now + TimeValue("0:00:01"))
        Loop

    Loop
    
    Dim url3 As String
    url3 = "http://urosys-web.juroinstruments.com/app/SelectJob1?jobid=" & jobId
    
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

