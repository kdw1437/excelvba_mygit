Attribute VB_Name = "ClassPostPrice"
Sub ClassPostPrice()
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("Market Data")
    
    Dim StartingPoint As Range
    Set StartingPoint = Ws.Range(Ws.Range("P2").value)
    
    Dim Table1Point As Range
    Set Table1Point = StartingPoint.Offset(3, 0)
    
    ' Equity table 다음에 위치하는 FX를 포함하는 cell을 찾는다.
    Dim fxRow As Range
    Set fxRow = Ws.Range(Table1Point.Offset(1, 0), Ws.Cells(Ws.Rows.Count, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)

    Dim PostPriceUpdater As New PostPriceUpdater
    With PostPriceUpdater
        Set .Worksheet = Ws
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

