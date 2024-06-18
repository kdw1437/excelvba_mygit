Attribute VB_Name = "ClassPostPrice"
Dim requestHandler As CAsyncRequestHandler
Sub ClassPostPrice()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    
    Dim StartingPoint As Range
    Set StartingPoint = ws.Range(ws.Range("P2").value)
    
    Dim Table1Point As Range
    Set Table1Point = StartingPoint.Offset(3, 0)
    
    ' Equity table ������ ��ġ�ϴ� FX�� �����ϴ� cell�� ã�´�.
    Dim fxRow As Range
    Set fxRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(ws.Rows.Count, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, Lookat:=xlWhole)

    Dim PostPriceUpdater As New PostPriceUpdater
    With PostPriceUpdater
        Set .Worksheet = ws
        Set .Table1Point = Table1Point
        Set .fxRow = fxRow
    End With
    
    Dim DataString As String
    DataString = PostPriceUpdater.GenerateJson2() ' JSON string�� �����.
    
    Debug.Print DataString
    DataString = URLEncode(DataString)
    
    Dim dataSetId As String
    dataSetId = ws.Range("O2").value
    
    Dim baseDt As String
    baseDt = Format(ws.Range("A2").value, "yyyymmdd")
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/savePrices?baseDt=" & baseDt & "&dataSetId=" & dataSetId
    
    ' JSON data�� POST request�� ������ ���� subroutine�� ȣ���Ѵ�.
    ' SendPostRequest DataString, url
    Set requestHandler = New CAsyncRequestHandler
    ' POST request�� ������.
    'SendPostRequest DataString, url
    requestHandler.SendPostRequestAsync DataString, url
End Sub

