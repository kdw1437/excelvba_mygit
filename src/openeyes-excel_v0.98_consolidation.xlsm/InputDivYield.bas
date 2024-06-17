Attribute VB_Name = "InputDivYield"
Sub InputDivYield()
    Dim divYieldUrlBuilder As UrlBuilder
    Set divYieldUrlBuilder = New UrlBuilder
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dividend")
    
    Dim baseDt As String
    baseDt = Format(ws.Range("A2").value, "yyyymmdd")
    
    'setter를 이용해서 UrlBuilder의 property를 적절하게 세팅해준다.
    divYieldUrlBuilder.baseurl = "http://localhost:8080/val/marketdata/"
    divYieldUrlBuilder.Version = "v1/"
    divYieldUrlBuilder.DataParameter = "selectDividends?"
    divYieldUrlBuilder.baseDt = "baseDt=" & baseDt & "&"
    divYieldUrlBuilder.DataIds = "dataIds=EUROSTOXX_C,KOSPI200_C,SPX_C"
    'divYieldUrlBuilder.DataIds = "dataIds=esd_c"
    '메서드 이용, return값이 full url.
    Dim divUrl As String
    divUrl = divYieldUrlBuilder.MakeUrl
    
    Debug.Print divUrl
    
    Dim JsonString As String
    JsonString = GetHttpResponseText2(divUrl)
    
    Dim JsonResponse As Object
    Set JsonResponse = JsonConverter.ParseJson(JsonString)
    
    If JsonResponse.Exists("code") Then
        If JsonResponse("code") = "ERROR" Then
            Dim errMsg As String
            errMsg = "Error: " & JsonResponse("message")
            MsgBox errMsg, vbCritical ' Display the error message in a message box
            Exit Sub
        
        ElseIf JsonResponse("code") = "SUCCESS" Then
            Dim divYields As Collection
            Set divYields = JsonResponse("response")("dividendYields")
        
            Dim divYieldUpdater As divYieldUpdater
            Set divYieldUpdater = New divYieldUpdater
            
            divYieldUpdater.SetWorksheet "Dividend"
            'divYieldUpdater.SetDivRange "A5:A7" '이 부분 하드코딩되어 있는데, 수정 필요할 시 수정
            
            Dim lastRow As Long
            lastRow = ws.Range("A5").End(xlDown).row
            Dim divRange As String
            divRange = "A5:A" & lastRow
            
            divYieldUpdater.SetDivRange divRange
            
            divYieldUpdater.SetDivData divYields
            
            'divYieldUpdater.UpdateWorksheet
            divYieldUpdater.UpdateWorksheetEfficient
        End If
    End If
    
End Sub
