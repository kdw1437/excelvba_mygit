Attribute VB_Name = "ConvertRangeToJson"
Option Explicit

Dim httpHandler As New clsXMLHTTPHandler


Sub ConvertRangeToJson()
    Dim PostQuoteUpdaterNew As New PostQuoteUpdaterNew ' 실제 클래스 이름 넣기
    Dim ws As Worksheet
    Dim rng As Range
    Dim jsonString As String

    ' worksheet 세팅
    
    Set PostQuoteUpdaterNew.Worksheet = ThisWorkbook.Sheets("Quote")
    Set ws = PostQuoteUpdaterNew.Worksheet
    
    ' range를 정의한다.
    ' A9이 가장 왼쪽 위의 cell이다.
    ' 실제 데이터에 따라서 range를 조정한다.
    'Set PostQuoteUpdaterNew.Range = PostQuoteUpdaterNew.Worksheet.Range("A9:Y13") (하드코딩 적으로 range를 정해줄 때)
    Dim lastCol As Range
    Dim lastRow As Range
    Dim endCell As Range
    
    Set lastCol = ws.Range("A9").End(xlToRight)
    Set lastRow = ws.Range("A9").End(xlDown)
    Set endCell = ws.Cells(lastRow.row, lastCol.Column)
    
    Set rng = ws.Range("A9", endCell)
    Set PostQuoteUpdaterNew.Range = rng
    
    ' range를 JSON으로 바꾼다.
    jsonString = PostQuoteUpdaterNew.makeJsonString2()

    ' JSONString 콘솔에 출력
    Debug.Print jsonString
    
    Dim dataSetId As String
    dataSetId = "official"
    
    Dim baseDt As String
    baseDt = Format(ws.Range("A2").value, "yyyymmdd")
    
    jsonString = URLEncode(jsonString)
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveQuoteIssueInfo?baseDt=" & baseDt & "&dataSetId=" & dataSetId
    
    
    
    ' POST request를 보내기 위해 method를 호출한다.
    httpHandler.SendPostRequest jsonString, url
    
 
    
    

End Sub



