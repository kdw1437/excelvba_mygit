Attribute VB_Name = "ClassPostQuote"
Sub ClassPostQuote()
    
    Dim quoteUpdater As postQuoteUpdater
    Set quoteUpdater = New postQuoteUpdater
    
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
    
    jsonString = URLEncode(jsonString)
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveQuoteIssueInfo?baseDt=20231228&dataSetId=TEST2"
    
    ' JSON data와 POST request를 보내기 위해 subroutine을 호출한다.
    SendPostRequest jsonString, url
End Sub

Sub ConvertRangeToJson()
    Dim postQuoteUpdaterNew As New postQuoteUpdaterNew ' Replace with the actual name of your class
    Dim ws As Worksheet
    Dim rng As Range
    Dim jsonString As String

    ' Set the worksheet
    
    Set postQuoteUpdaterNew.Worksheet = ThisWorkbook.Sheets("Quote_2")
    
    ' Define the range that you want to convert
    ' Assuming A9 is the top left cell of your data and the range extends to column R with 4 data rows
    ' You would adjust the range according to your actual data
    Set postQuoteUpdaterNew.Range = postQuoteUpdaterNew.Worksheet.Range("A9:Y13")

    ' Convert the range to JSON
    jsonString = postQuoteUpdaterNew.ToJsonString

    ' Do something with the JSON string, for example, output to Immediate Window
    Debug.Print jsonString
End Sub

