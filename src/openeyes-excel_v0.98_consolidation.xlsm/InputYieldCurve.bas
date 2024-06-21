Attribute VB_Name = "InputYieldCurve"
Sub InputYieldCurve()
    Dim YCUrlBuilder As UrlBuilder
    Set YCUrlBuilder = New UrlBuilder
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    
    Dim baseDt As String
    baseDt = Format(ws.Range("A2").value, "yyyymmdd")
    
    YCUrlBuilder.baseurl = "http://localhost:8080/val/marketdata/"
    YCUrlBuilder.Version = "v1/"
    YCUrlBuilder.DataParameter = "yieldcurves?"
    'YCUrlBuilder.baseDt = "baseDt=20231228&"
    YCUrlBuilder.baseDt = "baseDt=" & baseDt & "&"
    'YCUrlBuilder.dataIds = "dataIds=KRWIRSZ,JPYIRSZ,EURIRSZ,HKDIRSZ,USDIRSZ"
    
    ' column A에서 yieldCurveCell을 찾는다.
    Dim yieldCurveCell As Range
    Set yieldCurveCell = ws.Columns("A").Find(What:="Yield Curve", LookIn:=xlValues, LookAt:=xlWhole)
    
    Dim dataIds As String
    If Not yieldCurveCell Is Nothing Then
        ' "Yield Curve" cell로 부터 2 row 밑에서 시작한다.
        Dim startCell As Range
        Set startCell = yieldCurveCell.Offset(2, 0)
        
        ' 빈 cell을 찾을 때 까지 2 column씩 이동한다.
        Dim currentCell As Range
        Set currentCell = startCell
        
        Do While Not IsEmpty(currentCell.value)
            If dataIds <> "" Then
                dataIds = dataIds & ","
            End If
            dataIds = dataIds & currentCell.value
            Set currentCell = currentCell.Offset(0, 2)
        Loop
    End If
    
    YCUrlBuilder.dataIds = "dataIds=" & dataIds
    
    Dim YCUrl As String
    YCUrl = YCUrlBuilder.MakeUrl
    
    Debug.Print YCUrl
    
    Dim jsonString As String
    jsonString = GetHttpResponseText2(YCUrl)
    
    Dim JsonResponse As Object
    Set JsonResponse = JsonConverter.ParseJson(jsonString)
    
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
                
'            Dim ws As Worksheet
'            Set ws = ThisWorkbook.Sheets("Market Data")
            
            With yieldCurveUpdater1
                Set .Worksheet = ws
                Set .yieldCurves = yieldCurves
                Set .CurrencyCell = ws.Range("A27:J27")
                .PopulateYieldCurveData
                        
            End With
        End If
    End If
    
    
End Sub
