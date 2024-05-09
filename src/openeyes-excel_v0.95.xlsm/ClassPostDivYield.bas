Attribute VB_Name = "ClassPostDivYield"
Sub ClassPostDivYield()
    
    Dim divYieldUpdater As PostDivYieldUpdater
    Set divYieldUpdater = New PostDivYieldUpdater
    
    Set divYieldUpdater.Worksheet = ThisWorkbook.Sheets("Dividend")
    Set divYieldUpdater.DivCell = divYieldUpdater.Worksheet.Range("F3")
    'divYieldUpdater.DataIdRange = "F5:F7"
    Set divYieldUpdater.DataIdRange = divYieldUpdater.Worksheet.Range("F5:F7")
    
    Dim JsonString As String
    JsonString = divYieldUpdater.GenerateJson2()
    
    Debug.Print JsonString
    
End Sub
