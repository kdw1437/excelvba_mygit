Attribute VB_Name = "ConvertRangeToJson"
Sub ConvertRangeToJson()
    Dim PostQuoteUpdaterNew As New PostQuoteUpdaterNew ' Replace with the actual name of your class
    Dim ws As Worksheet
    Dim rng As Range
    Dim JsonString As String

    ' Set the worksheet
    
    Set PostQuoteUpdaterNew.Worksheet = ThisWorkbook.Sheets("Quote")
    Set ws = PostQuoteUpdaterNew.Worksheet
    
    ' Define the range that you want to convert
    ' Assuming A9 is the top left cell of your data and the range extends to column R with 4 data rows
    ' You would adjust the range according to your actual data
    'Set PostQuoteUpdaterNew.Range = PostQuoteUpdaterNew.Worksheet.Range("A9:Y13")
    Dim lastCol As Range
    Dim lastRow As Range
    Dim endCell As Range
    
    Set lastCol = ws.Range("A9").End(xlToRight)
    Set lastRow = ws.Range("A9").End(xlDown)
    Set endCell = ws.Cells(lastRow.row, lastCol.Column)
    
    Set rng = ws.Range("A9", endCell)
    Set PostQuoteUpdaterNew.Range = rng
    
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



