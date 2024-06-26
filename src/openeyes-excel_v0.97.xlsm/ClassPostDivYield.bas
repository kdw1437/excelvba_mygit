Attribute VB_Name = "ClassPostDivYield"
Dim requestHandler As CAsyncRequestHandler

Sub ClassPostDivYield()
    
    Dim divYieldUpdater As PostDivYieldUpdater
    Set divYieldUpdater = New PostDivYieldUpdater
    
    Set divYieldUpdater.Worksheet = ThisWorkbook.Sheets("Dividend")
    Set divYieldUpdater.DivCell = divYieldUpdater.Worksheet.Range("F3")
    
    Set divYieldUpdater.DataIdRange = divYieldUpdater.Worksheet.Range(divYieldUpdater.DivCell.Offset(2, 0), divYieldUpdater.DivCell.Offset(2, 0).End(xlDown))
    
    'divYieldUpdater.DataIdRange = "F5:F7"
    'Set divYieldUpdater.DataIdRange = divYieldUpdater.Worksheet.Range("F5:F7")
    
    Dim JsonString As String
    JsonString = divYieldUpdater.GenerateJson2()
    
    Debug.Print JsonString
    
    JsonString = URLEncode(JsonString)
    
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("Dividend")
    
    Dim dataSetId As String
    dataSetId = Ws.Range("F2").value
    
    Dim baseDt As String
    baseDt = Format(Ws.Range("A2").value, "yyyymmdd")
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveDividends?baseDt=" & baseDt & "&dataSetId=" & dataSetId

    'SendPostRequest JsonString, url
    
    Set requestHandler = New CAsyncRequestHandler
    ' POST request�� ������.
    requestHandler.SendPostRequestAsync JsonString, url
End Sub
