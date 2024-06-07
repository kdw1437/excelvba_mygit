Attribute VB_Name = "ClassPostQuote"
Sub ConvertRangeToJson()
    Dim PostQuoteUpdaterNew As New PostQuoteUpdaterNew ' Replace with the actual name of your class
    Dim ws As Worksheet
    Dim rng As Range
    Dim JsonString As String

    ' worksheet를 세팅한다.
    
    Set PostQuoteUpdaterNew.Worksheet = ThisWorkbook.Sheets("Quote_2")
    Set ws = PostQuoteUpdaterNew.Worksheet
    
    ' 변경하고자 하는 range를 정의한다.
    ' A9cell이 첫번째 데이터의 가장 왼쪽 cell이다.
    ' 실제 데이터에 따라 range를 수정한다.
    'Set PostQuoteUpdaterNew.Range = PostQuoteUpdaterNew.Worksheet.Range("A9:Y13")
    Dim lastCol As Range
    Dim lastRow As Range
    Dim endCell As Range
    
    Set lastCol = ws.Range("A9").End(xlToRight)
    Set lastRow = ws.Range("A9").End(xlDown)
    Set endCell = ws.Cells(lastRow.row, lastCol.Column)
    
    Set rng = ws.Range("A9", endCell)
    Set PostQuoteUpdaterNew.Range = rng
    
    ' range안의 데이터를 JSON타입 데이터로 바꾼다.
    JsonString = PostQuoteUpdaterNew.makeJsonString2()

    ' JsonString을 console에 표시한다.
    Debug.Print JsonString
    
    JsonString = URLEncode(JsonString)
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveQuoteIssueInfo?baseDt=20231228&dataSetId=TEST2"
    
    ' JSON data와 POST request를 보내기 위해 subroutine을 호출한다.
    SendPostRequest JsonString, url
End Sub

