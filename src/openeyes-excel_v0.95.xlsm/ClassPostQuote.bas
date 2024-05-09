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
    
    Dim JsonString As String
    JsonString = quoteUpdater.makeJsonString()
    
    Debug.Print JsonString
    
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
    
    Dim JsonString As String
    JsonString = quoteUpdater.makeJsonString()
    
    Debug.Print JsonString
    
    
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
    
    Dim JsonString As String
    JsonString = quoteUpdater.makeJsonString()
    
    Debug.Print JsonString
    
    JsonString = URLEncode(JsonString)
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveQuoteIssueInfo?baseDt=20231228&dataSetId=TEST2"
    
    ' JSON data와 POST request를 보내기 위해 subroutine을 호출한다.
    SendPostRequest JsonString, url
End Sub

Sub ConvertRangeToJson()
    Dim PostQuoteUpdaterNew As New PostQuoteUpdaterNew ' Replace with the actual name of your class
    Dim ws As Worksheet
    Dim rng As Range
    Dim JsonString As String

    ' Set the worksheet
    
    Set PostQuoteUpdaterNew.Worksheet = ThisWorkbook.Sheets("Quote_2")
    
    ' Define the range that you want to convert
    ' Assuming A9 is the top left cell of your data and the range extends to column R with 4 data rows
    ' You would adjust the range according to your actual data
    Set PostQuoteUpdaterNew.Range = PostQuoteUpdaterNew.Worksheet.Range("A9:Y13")

    ' Convert the range to JSON
    JsonString = PostQuoteUpdaterNew.makeJsonString2()

    ' Do something with the JSON string, for example, output to Immediate Window
    Debug.Print JsonString
    
    JsonString = URLEncode(JsonString)
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveQuoteIssueInfo?baseDt=20231228&dataSetId=TEST2"
    
    ' JSON data와 POST request를 보내기 위해 subroutine을 호출한다.
    SendPostRequest JsonString, url
End Sub

