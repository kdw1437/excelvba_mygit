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
    
    '메서드 이용, return값이 full url.
    Dim divUrl As String
    divUrl = divYieldUrlBuilder.MakeUrl
    
    Debug.Print divUrl
    
    Dim JsonString As String
    JsonString = GetHttpResponseText(divUrl)
    
    Dim JsonResponse As Object
    Set JsonResponse = JsonConverter.ParseJson(JsonString)
    
    Dim divYields As Collection
    Set divYields = JsonResponse("response")("dividendYields")

    Dim divYieldUpdater As divYieldUpdater
    Set divYieldUpdater = New divYieldUpdater
    
    divYieldUpdater.SetWorksheet "Dividend"
    divYieldUpdater.SetDivRange "A5:A7"
    divYieldUpdater.SetDivData divYields
    
    'divYieldUpdater.UpdateWorksheet
    divYieldUpdater.UpdateWorksheetEfficient
    
    
End Sub
