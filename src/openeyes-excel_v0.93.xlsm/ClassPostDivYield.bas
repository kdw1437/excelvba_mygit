Attribute VB_Name = "ClassPostDivYield"
Sub ClassPostDivYield()
    
    Dim divYieldUpdater As postDivYieldUpdater
    Set divYieldUpdater = New postDivYieldUpdater
    
    Set divYieldUpdater.Worksheet = ThisWorkbook.Sheets("Dividend")
    Set divYieldUpdater.DivCell = divYieldUpdater.Worksheet.Range("F3")
    'divYieldUpdater.DataIdRange = "F5:F7"
    Set divYieldUpdater.DataIdRange = divYieldUpdater.Worksheet.Range("F5:F7")
    
    Dim jsonString As String
    jsonString = divYieldUpdater.GenerateJson()
    
    Debug.Print jsonString
    
End Sub
