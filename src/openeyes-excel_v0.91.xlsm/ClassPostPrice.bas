Attribute VB_Name = "ClassPostPrice"
Sub ClassPostPrice()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    
    Dim StartingPoint As Range
    Set StartingPoint = ws.Range(ws.Range("P2").Value)
    
    Dim Table1Point As Range
    Set Table1Point = StartingPoint.Offset(3, 0)
    
    ' Equity table 다음에 위치하는 FX를 포함하는 cell을 찾는다.
    Dim fxRow As Range
    Set fxRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(ws.Rows.Count, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)

    Dim PostPriceUpdater As New PostPriceUpdater
    With PostPriceUpdater
        Set .Worksheet = ws
        Set .Table1Point = Table1Point
        Set .fxRow = fxRow
    End With
    
    Dim DataString As String
    DataString = PostPriceUpdater.GenerateJson() ' JSON string을 만든다.
    
    Debug.Print DataString
    DataString = URLEncode(DataString)
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/prices?baseDt=20231228&dataSetId=TEST11"
    
    ' JSON data와 POST request를 보내기 위해 subroutine을 호출한다.
    SendPostRequest DataString, url
End Sub

