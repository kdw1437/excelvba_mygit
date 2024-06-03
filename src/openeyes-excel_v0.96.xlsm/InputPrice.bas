Attribute VB_Name = "InputPrice"
Sub InputPrice()
    Dim priceUrlBuilder As UrlBuilder
    Set priceUrlBuilder = New UrlBuilder
    
    'setter를 이용해서 UrlBuilder의 property를 적절하게 세팅해준다.
    priceUrlBuilder.baseURL = "http://localhost:8080/val/marketdata/"
    priceUrlBuilder.Version = "v1/"
    priceUrlBuilder.DataParameter = "prices?"
    priceUrlBuilder.baseDt = "baseDt=20231228&"
    priceUrlBuilder.DataIds = "dataIds=KOSPI200,SPX,N225,EUROSTOXX,HSCEI,HSI,KR7035420009"
    
    '메서드 이용, return값이 full url.
    Dim priceUrl As String
    priceUrl = priceUrlBuilder.MakeUrl
    
    Debug.Print priceUrl
    
    Dim JsonString As String
    JsonString = GetHttpResponseText(priceUrl)
    
    Dim JsonResponse As Object
    Set JsonResponse = JsonConverter.ParseJson(JsonString)
    
    If JsonResponse.Exists("code") Then
        If JsonResponse("code") = "ERROR" Then
            Dim errMsg As String
            errMsg = "Error: " & JsonResponse("message")
            MsgBox errMsg, vbCritical ' Display the error message in a message box
            Exit Sub
        
        ElseIf JsonResponse("code") = "SUCCESS" Then
            Dim prices As Collection
            Set prices = JsonResponse("response")("prices")
            
            Dim priceMarketDataUpdater As MarketDataUpdater
            Set priceMarketDataUpdater = New MarketDataUpdater
            
            Set priceMarketDataUpdater.Worksheet = ThisWorkbook.Sheets("Market Data")
            Set priceMarketDataUpdater.PricesCollection = prices
            
            'priceMarketDataUpdater.UpdatePrices
            'priceMarketDataUpdater.UpdatePrices2
            priceMarketDataUpdater.UpdatePricesOptimized
            
            
        End If
        
    End If
End Sub
