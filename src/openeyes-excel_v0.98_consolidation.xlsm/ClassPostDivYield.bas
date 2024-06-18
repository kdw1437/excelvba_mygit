Attribute VB_Name = "ClassPostDivYield"
Dim requestHandler As CAsyncRequestHandler

Sub ClassPostDivYield()
    
    Dim divYieldUpdater As PostDivYieldUpdater
    Set divYieldUpdater = New PostDivYieldUpdater
    
    Set divYieldUpdater.Worksheet = ThisWorkbook.Sheets("Dividend")
    Set divYieldUpdater.DivCell = divYieldUpdater.Worksheet.Range("F3")
    
    Set divYieldUpdater.dataIdRange = divYieldUpdater.Worksheet.Range(divYieldUpdater.DivCell.Offset(2, 0), divYieldUpdater.DivCell.Offset(2, 0).End(xlDown))
    
    'divYieldUpdater.DataIdRange = "F5:F7"
    'Set divYieldUpdater.DataIdRange = divYieldUpdater.Worksheet.Range("F5:F7")
    
    Dim jsonString As String
    jsonString = divYieldUpdater.GenerateJson2()
    
    Debug.Print jsonString
    
    jsonString = URLEncode(jsonString)
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dividend")
    
    Dim dataSetId As String
    dataSetId = ws.Range("F2").value
    
    Dim baseDt As String
    baseDt = Format(ws.Range("A2").value, "yyyymmdd")
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveDividends?baseDt=" & baseDt & "&dataSetId=" & dataSetId

    'SendPostRequest JsonString, url
    
    Set requestHandler = New CAsyncRequestHandler
    ' POST request를 보낸다.
    requestHandler.SendPostRequestAsync jsonString, url
End Sub
