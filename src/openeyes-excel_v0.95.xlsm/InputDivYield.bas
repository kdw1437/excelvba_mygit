Attribute VB_Name = "InputDivYield"
Sub InputPrice()
    Dim divYieldUrlBuilder As UrlBuilder
    Set divYieldUrlBuilder = New UrlBuilder
    
    'setter를 이용해서 UrlBuilder의 property를 적절하게 세팅해준다.
    divYieldUrlBuilder.baseURL = "http://localhost:8080/val/marketdata/"
    divYieldUrlBuilder.Version = "v1/"
    divYieldUrlBuilder.DataParameter = "selectDividends?"
    divYieldUrlBuilder.baseDt = "baseDt=20211111&"
    divYieldUrlBuilder.DataIds = "dataIds=EUROSTOXX_C,KOSPI200_C,SPX_C"
    'divYieldUrlBuilder.DataIds = "dataIds=esd_c"
    '메서드 이용, return값이 full url.
    Dim divUrl As String
    divUrl = divYieldUrlBuilder.MakeUrl
    
    Debug.Print divUrl
    
    Dim JsonString As String
    JsonString = GetHttpResponseText(divUrl)
    
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
            divYieldUpdater.SetDivRange "A5:A7" '이 부분 하드코딩되어 있는데, 수정 필요할 시 수정
            divYieldUpdater.SetDivData divYields
            
            'divYieldUpdater.UpdateWorksheet
            divYieldUpdater.UpdateWorksheetEfficient
        End If
    End If
    
End Sub
