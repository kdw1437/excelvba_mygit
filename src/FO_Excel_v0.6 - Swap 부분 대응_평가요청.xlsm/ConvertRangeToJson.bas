Attribute VB_Name = "ConvertRangeToJson"
Option Explicit

Dim httpHandler As New clsXMLHTTPHandler


Sub ConvertRangeToJson()
    Dim PostQuoteUpdaterNew As New PostQuoteUpdaterNew ' ���� Ŭ���� �̸� �ֱ�
    Dim ws As Worksheet
    Dim rng As Range
    Dim JsonString As String

    ' worksheet ����
    
    Set PostQuoteUpdaterNew.Worksheet = ThisWorkbook.Sheets("Quote")
    Set ws = PostQuoteUpdaterNew.Worksheet
    
    ' range�� �����Ѵ�.
    ' A9�� ���� ���� ���� cell�̴�.
    ' ���� �����Ϳ� ���� range�� �����Ѵ�.
    'Set PostQuoteUpdaterNew.Range = PostQuoteUpdaterNew.Worksheet.Range("A9:Y13") (�ϵ��ڵ� ������ range�� ������ ��)
    Dim lastCol As Range
    Dim lastRow As Range
    Dim endCell As Range
    
    Set lastCol = ws.Range("A9").End(xlToRight)
    Set lastRow = ws.Range("A9").End(xlDown)
    Set endCell = ws.Cells(lastRow.row, lastCol.Column)
    
    Set rng = ws.Range("A9", endCell)
    Set PostQuoteUpdaterNew.Range = rng
    
    ' range�� JSON���� �ٲ۴�.
    JsonString = PostQuoteUpdaterNew.makeJsonString2()

    ' JSONString �ֿܼ� ���
    Debug.Print JsonString
    
    JsonString = URLEncode(JsonString)
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveQuoteIssueInfo?baseDt=20231228&dataSetId=TEST2"
    
    
    
    ' POST request�� ������ ���� method�� ȣ���Ѵ�.
    httpHandler.SendPostRequest JsonString, url
    
 
    
    

End Sub



