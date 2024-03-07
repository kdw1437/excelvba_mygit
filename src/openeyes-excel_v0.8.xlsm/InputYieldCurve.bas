Attribute VB_Name = "InputYieldCurve"
Sub InputYieldCurve()
    Dim YCUrlBuilder As UrlBuilder
    Set YCUrlBuilder = New UrlBuilder
    
    YCUrlBuilder.baseURL = "http://localhost:8080/val/marketdata/"
    YCUrlBuilder.Version = "v1/"
    YCUrlBuilder.DataParameter = "yieldcurves?"
    YCUrlBuilder.baseDt = "baseDt=20231228&"
    YCUrlBuilder.DataIds = "dataIds=KRWIRSZ,JPYIRSZ,EURIRSZ,HKDIRSZ,USDIRSZ"
    
    Dim YCUrl As String
    YCUrl = YCUrlBuilder.MakeUrl
    
    Debug.Print YCUrl
    
    Dim JsonResponse As Object
    Set JsonResponse = GetJsonResponse(YCUrl)
    Dim YieldCurves As Collection
    Set YieldCurves = JsonResponse("response")("yieldCurves")
    
    Dim yieldCurve As Variant
    Set yieldCurve = YieldCurves(1)
    
    Dim FX As String
    FX = yieldCurve("dataId")
    
    Dim yieldCurveUpdater1 As YieldCurveUpdater
    Set yieldCurveUpdater1 = New YieldCurveUpdater
        
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    
    With yieldCurveUpdater1
        Set .Worksheet = ws
        Set .YieldCurves = YieldCurves
        Set .CurrencyCell = ws.Range("A27:J27")
        .PopulateYieldCurveData
                
    End With
        
    
    
    
End Sub
