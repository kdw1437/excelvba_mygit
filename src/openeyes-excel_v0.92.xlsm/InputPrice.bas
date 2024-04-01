Attribute VB_Name = "InputPrice"
Sub InputPrice()
    Dim priceUrlBuilder As UrlBuilder
    Set priceUrlBuilder = New UrlBuilder
    
    'setter�� �̿��ؼ� UrlBuilder�� property�� �����ϰ� �������ش�.
    priceUrlBuilder.baseURL = "http://localhost:8080/val/marketdata/"
    priceUrlBuilder.Version = "v1/"
    priceUrlBuilder.DataParameter = "prices?"
    priceUrlBuilder.baseDt = "baseDt=20231228&"
    priceUrlBuilder.DataIds = "dataIds=KOSPI200,SPX,N225,EUROSTOXX,HSCEI,HSI,KR7035420009"
    
    '�޼��� �̿�, return���� full url.
    Dim priceUrl As String
    priceUrl = priceUrlBuilder.MakeUrl
    
    Debug.Print priceUrl
    
    Dim jsonString As String
    jsonString = GetHttpResponseText(priceUrl)
    
    Dim JsonResponse As Object
    Set JsonResponse = JsonConverter.ParseJson(jsonString)
    
    Dim prices As Collection
    Set prices = JsonResponse("response")("prices")
    
    Dim priceMarketDataUpdater As MarketDataUpdater
    Set priceMarketDataUpdater = New MarketDataUpdater
    
    Set priceMarketDataUpdater.Worksheet = ThisWorkbook.Sheets("Market Data")
    Set priceMarketDataUpdater.PricesCollection = prices
    
    priceMarketDataUpdater.UpdatePrices
    
       
End Sub