Attribute VB_Name = "ClassPostDivYield"
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

    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveDividends?baseDt=20231228&dataSetId=TEST2"

    ' POST request를 보낸다.
    SendPostRequest JsonString, url
End Sub
