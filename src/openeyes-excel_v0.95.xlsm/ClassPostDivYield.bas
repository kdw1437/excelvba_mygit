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
    
End Sub
