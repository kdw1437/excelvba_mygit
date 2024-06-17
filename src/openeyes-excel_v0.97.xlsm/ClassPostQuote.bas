Attribute VB_Name = "ClassPostQuote"
Sub ConvertRangeToJson()
    Dim PostQuoteUpdaterNew As New PostQuoteUpdaterNew ' Replace with the actual name of your class
    Dim ws As Worksheet
    Dim rng As Range
    Dim JsonString As String

    ' worksheet�� �����Ѵ�.
    
    Set PostQuoteUpdaterNew.Worksheet = ThisWorkbook.Sheets("Quote_2")
    Set ws = PostQuoteUpdaterNew.Worksheet
    
    ' �����ϰ��� �ϴ� range�� �����Ѵ�.
    ' A9cell�� ù��° �������� ���� ���� cell�̴�.
    ' ���� �����Ϳ� ���� range�� �����Ѵ�.
    'Set PostQuoteUpdaterNew.Range = PostQuoteUpdaterNew.Worksheet.Range("A9:Y13")
    Dim lastCol As Range
    Dim lastRow As Range
    Dim endCell As Range
    
    Set lastCol = ws.Range("A9").End(xlToRight)
    Set lastRow = ws.Range("A9").End(xlDown)
    Set endCell = ws.Cells(lastRow.row, lastCol.Column)
    
    Set rng = ws.Range("A9", endCell)
    Set PostQuoteUpdaterNew.Range = rng
    
    ' range���� �����͸� JSONŸ�� �����ͷ� �ٲ۴�.
    JsonString = PostQuoteUpdaterNew.makeJsonString2()

    ' JsonString�� console�� ǥ���Ѵ�.
    Debug.Print JsonString
    
    JsonString = URLEncode(JsonString)
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveQuoteIssueInfo?baseDt=20231228&dataSetId=TEST2"
    
    ' JSON data�� POST request�� ������ ���� subroutine�� ȣ���Ѵ�.
    SendPostRequest JsonString, url
End Sub

