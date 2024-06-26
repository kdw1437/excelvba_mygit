Attribute VB_Name = "Testandcomplete"
Sub SendJSON()
    Dim JsonString As String
    Dim xmlhttp As Object
    
    ' code에서 JSON string을 바로 정의한다.
    JsonString = StrConv("test=test1", vbFromUnicode)
    
    ' 새로운 XML HTTP request를 만든다.
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' 평가를 요청하는 url
    Dim url As String
    url = "http://urosys-web.juroinstruments.com/app/createValWebJob"
    
    ' post를 위한 xmlhttp의 메소드 Open 호출
    xmlhttp.Open "POST", url, False
    
    ' Set the request content-type header to application/json
    xmlhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    
    ' Send the request with the JSON string
    xmlhttp.Send "officeCd=BO&name=TEST4&valDate=20231228&valTypeCode=P&greekLevel=&contextIds=BO&dataSetIds=Test_4,official&simId=&priority=4&itemCodes=ELS3588"
    
    ' Check the status of the request
    If xmlhttp.Status = 200 Then


        Dim response As String
        response = xmlhttp.ResponseText
        
        Dim jobId As String
        Dim ParsedJson As Dictionary
        
        Set ParsedJson = JsonConverter.ParseJson(response)
        jobId = ParsedJson("jobId")
        
        Dim url2 As String
        url2 = "http://urosys-web.juroinstruments.com/app/selectValJob?jobId=" & jobId
        
        Dim startTime As Double
        startTime = Timer
        
        Do
            Dim xmlhttp2 As Object
            Set xmlhttp2 = CreateObject("WinHttp.WinHttpRequest.5.1")
            
            xmlhttp2.Open "GET", url2, False
            xmlhttp2.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            xmlhttp2.Send
            
            response = xmlhttp2.ResponseText
            Set ParsedJson = JsonConverter.ParseJson(response)
            
            If ParsedJson("jobStateCodeNm") = "완료" Then
                MsgBox "평가 작업이 완료되었습니다."
                Dim url3 As String
                url3 = "http://urosys-web.juroinstruments.com/app/SelectJob1?jobid=" & jobId
                
                Dim httpRequest As Object
                Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
                With httpRequest
                    .Open "GET", url3, False
                    .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                    .Send
                    JsonString = .ResponseText
                End With
                
                Debug.Print JsonString
                
                Set jsonResponse = JsonConverter.ParseJson(JsonString)
                Set ws = ThisWorkbook.Worksheets("Sheet1")


                ' Starting row
                Dim row As Integer
                row = 1
            
                ' Column numbers
                Const COL_ITEM_CD As Integer = 1
                Const COL_PRICE As Integer = 2
            
                ' Iterate through the JSON array and write data to the worksheet
                For Each Item In jsonResponse("selectjob1")
                    ws.Cells(row, COL_ITEM_CD).Value = Item("itemCd")
                    ws.Cells(row, COL_PRICE).Value = Item("price")
                    row = row + 1
                Next Item
                Exit Do
                
            ElseIf Timer - startTime > 10 Then
                MsgBox "평가 작업이 완료되지 않았습니다."
                Exit Do
            End If
            
            DoEvents
        Loop
        Set xmlhttp2 = Nothing
    Else
        ' If the request failed, output the status
        MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.StatusText
    End If
    ' Clean up
    Set xmlhttp = Nothing
    Set ParsedJson = Nothing
End Sub

'평가작업'

