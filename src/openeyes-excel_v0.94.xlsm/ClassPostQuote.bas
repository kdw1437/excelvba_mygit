Attribute VB_Name = "ClassPostQuote"
Sub ClassPostQuote()
    
    Dim quoteUpdater As PostQuoteUpdater
    Set quoteUpdater = New PostQuoteUpdater
    
    Set quoteUpdater.Worksheet = ThisWorkbook.Sheets("Quote")
    
    'divYieldUpdater.DataIdRange = "F5:F7"
    
    Dim startCell As Range
    Dim lastColumnCell As Range
    Dim lastRowCell As Range
    
    Set startCell = ThisWorkbook.Sheets("Quote").Range("A9")
    Set lastColumnCell = startCell.End(xlToRight)
    Set lastRowCell = startCell.End(xlDown)
    
    Set quoteUpdater.Range = quoteUpdater.Worksheet.Range(startCell, quoteUpdater.Worksheet.Cells(lastRowCell.row, lastColumnCell.Column))
    
    
    'Set quoteUpdater.Range = quoteUpdater.Worksheet.Range("A9:AA11")
    
    Dim jsonString As String
    jsonString = quoteUpdater.makeJsonString()
    
    Debug.Print jsonString
    
End Sub

Sub ClassPostQuoteNumeric()

    Dim quoteUpdater As PostQuoteNumeric
    Set quoteUpdater = New PostQuoteNumeric
    
    Set quoteUpdater.Worksheet = ThisWorkbook.Sheets("Quote")
    
    Dim startCell As Range
    Dim lastColumnCell As Range
    Dim lastRowCell As Range
    
    Set startCell = ThisWorkbook.Sheets("Quote").Range("A9")
    Set lastColumnCell = startCell.End(xlToRight)
    Set lastRowCell = startCell.End(xlDown)
    
    Set quoteUpdater.Range = quoteUpdater.Worksheet.Range(startCell, quoteUpdater.Worksheet.Cells(lastRowCell.row, lastColumnCell.Column))
    
    'Set quoteUpdater.Range = quoteUpdater.Worksheet.Range("A9:AA11")
    
    Dim jsonString As String
    jsonString = quoteUpdater.makeJsonString()
    
    Debug.Print jsonString
    
    
End Sub

Sub ClassPostQuoteRecent()

    Dim quoteUpdater As PostQuoteRecent
    Set quoteUpdater = New PostQuoteRecent
    
    Set quoteUpdater.Worksheet = ThisWorkbook.Sheets("Quote")
    
    Dim startCell As Range
    Dim lastColumnCell As Range
    Dim lastRowCell As Range
    
    Set startCell = ThisWorkbook.Sheets("Quote").Range("A9")
    Set lastColumnCell = startCell.End(xlToRight)
    Set lastRowCell = startCell.End(xlDown)
    
    Set quoteUpdater.Range = quoteUpdater.Worksheet.Range(startCell, quoteUpdater.Worksheet.Cells(lastRowCell.row, lastColumnCell.Column))
    
    'Set quoteUpdater.Range = quoteUpdater.Worksheet.Range("A9:AA11")
    
    Dim jsonString As String
    jsonString = quoteUpdater.makeJsonString()
    
    Debug.Print jsonString
    
    
End Sub


