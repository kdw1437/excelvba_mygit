Attribute VB_Name = "Testandcomplete"
Sub SendJSON()
    Dim JsonString As String
    Dim xmlhttp As Object
    
    ' JSON string을 코드에 직접적으로 정의한다.
    'JsonString = StrConv("test=test1", vbFromUnicode)
    
    ' 새로운 XML HTTP request를 만든다.
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' request를 보내는 URL
    Dim url As String
    url = "http://urosys-web.juroinstruments.com/app/createValWebJob"
    
    ' POST method로서 HTTP request를 연다.
    xmlhttp.Open "POST", url, False
    
    ' application/json으로 request content-type header를 설정한다.
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    ' JSON string(실제로는 query string format)으로 request를 보낸다. json data type으로 보내려면, Content-type을 application/json으로 설정해야 한다.
    xmlhttp.Send "officeCd=BO&name=TEST4&valDate=20231228&valTypeCode=P&greekLevel=&contextIds=BO&dataSetIds=Test_4,official&simId=&priority=4&itemCodes=ELS3588"
    
    ' request의 상태를 확인한다.
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
            xmlhttp2.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
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
                    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
                    .Send
                    JsonString = .ResponseText
                End With
                
                Debug.Print JsonString
                
                Set jsonResponse = JsonConverter.ParseJson(JsonString)
                Set ws = ThisWorkbook.Worksheets("Sheet1")


                ' 시작 row
                Dim row As Integer
                row = 1
            
                ' column number
                Const COL_ITEM_CD As Integer = 1
                Const COL_PRICE As Integer = 2
            
                ' JSON array를 iteration하며 data를 worksheet에 넣는다.
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
        ' request가 실패하면, status를 출력한다.
        MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.StatusText
    End If
    ' 가비지 컬렉션
    Set xmlhttp = Nothing
    Set ParsedJson = Nothing
End Sub

'평가작업'

