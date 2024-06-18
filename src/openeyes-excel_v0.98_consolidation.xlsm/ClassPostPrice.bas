Attribute VB_Name = "ClassPostPrice"
Dim requestHandler As CAsyncRequestHandler
Sub ClassPostPrice()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    
    Dim StartingPoint As Range
    Set StartingPoint = ws.Range(ws.Range("P2").value)
    
    Dim Table1Point As Range
    Set Table1Point = StartingPoint.Offset(3, 0)
    
    ' Equity table 다음에 위치하는 FX를 포함하는 cell을 찾는다.
    Dim fxRow As Range
    Set fxRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(ws.Rows.Count, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, Lookat:=xlWhole)

    Dim PostPriceUpdater As New PostPriceUpdater
    With PostPriceUpdater
        Set .Worksheet = ws
        Set .Table1Point = Table1Point
        Set .fxRow = fxRow
    End With
    
    Dim DataString As String
    DataString = PostPriceUpdater.GenerateJson2() ' JSON string을 만든다.
    
    Debug.Print DataString
    DataString = URLEncode(DataString)
    
    Dim dataSetId As String
    dataSetId = ws.Range("O2").value
    
    Dim baseDt As String
    baseDt = Format(ws.Range("A2").value, "yyyymmdd")
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/savePrices?baseDt=" & baseDt & "&dataSetId=" & dataSetId
    
    ' JSON data와 POST request를 보내기 위해 subroutine을 호출한다.
    ' SendPostRequest DataString, url
    Set requestHandler = New CAsyncRequestHandler
    ' POST request를 보낸다.
    'SendPostRequest DataString, url
    requestHandler.SendPostRequestAsync DataString, url
End Sub

