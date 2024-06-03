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
    
    If JsonResponse.Exists("code") Then
        If JsonResponse("code") = "ERROR" Then
            Dim errMsg As String
            errMsg = "Error: " & JsonResponse("message")
            MsgBox errMsg, vbCritical ' Display the error message in a message box
            Exit Sub
    
        ElseIf JsonResponse("code") = "SUCCESS" Then
            Dim yieldCurves As Collection
            Set yieldCurves = JsonResponse("response")("yieldCurves")
            
            Dim yieldCurve As Variant
            Set yieldCurve = yieldCurves(1)
            
            Dim FX As String
            FX = yieldCurve("dataId")
            
            Dim yieldCurveUpdater1 As YieldCurveUpdater
            Set yieldCurveUpdater1 = New YieldCurveUpdater
                
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Sheets("Market Data")
            
            With yieldCurveUpdater1
                Set .Worksheet = ws
                Set .yieldCurves = yieldCurves
                Set .CurrencyCell = ws.Range("A27:J27")
                .PopulateYieldCurveData
                        
            End With
        End If
    End If
    
    
End Sub
