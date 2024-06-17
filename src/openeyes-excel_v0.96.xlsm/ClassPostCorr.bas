Attribute VB_Name = "ClassPostCorr"
Sub PrintJsonString()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")

    Dim equityCell As Range
    Set equityCell = ws.Columns("M:M").Find(What:="Equity", LookIn:=xlValues, LookAt:=xlWhole)

    ' postCorrUpdater의 새로운 인스턴스 생성
    Dim postCorrUpdater As New postCorrUpdater
    With postCorrUpdater
        ' vertical range를 설정한다. (Equity cell로 부터 특정 범위에 있는)
        Set .VerticalRange = equityCell.Offset(4, 0).Resize(ws.Range(equityCell.Offset(4, 0), ws.Cells(ws.Rows.Count, equityCell.Column)).End(xlDown).row - equityCell.Offset(4, 0).row + 1)
        
        ' horizontal range를 설정한다. (Equity cell로 부터 특정 범위에 있는)
        Set .HorizontalRange = equityCell.Offset(3, 2).Resize(, ws.Range(equityCell.Offset(3, 2), ws.Cells(equityCell.Offset(3, 2).row, ws.Columns.Count)).End(xlToRight).Column - equityCell.Offset(3, 2).Column + 1)
        
        
        'Debug.Print .CorrJson()
    End With
    Dim DataString As String
    DataString = postCorrUpdater.CorrJsonv()
    Debug.Print DataString
    DataString = URLEncode(DataString)
    Dim url As String
    
    Dim dataSetId As String
    dataSetId = ws.Range("O2").value
    'url = "http://localhost:8080/val/marketdata/v1/corrs?baseDt=20231228&dataSetId=TEST11&matrixId=CORR"
    'url = "http://localhost:8080/val/marketdata/v1/saveCorrs?baseDt=20231228&dataSetId=TEST15&matrixId=CORR"
    url = "http://localhost:8080/val/marketdata/v1/saveCorrs?baseDt=20231228&dataSetId=" & dataSetId & "&matrixId=CORR"
    ' JSON data와 POST request를 보내기 위해 subroutine을 호출한다.
    SendPostRequest DataString, url
End Sub

Sub PrintJsonString2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    
    Dim FXCell As Range
    Set FXCell = ws.Columns("M:M").Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)

    ' postCorrUpdater의 instance를 생성한다.
    Dim postCorrUpdater As New postCorrUpdater
    With postCorrUpdater
        ' VerticalRange property의 값을 부여한다. (setter이용)
        Set .VerticalRange = FXCell.Offset(4, 0).Resize(ws.Range(FXCell.Offset(4, 0), ws.Cells(ws.Rows.Count, FXCell.Column)).End(xlDown).row - FXCell.Offset(4, 0).row + 1)
        
        'HorizontalRange property의 값을 부여한다. (setter이용)
        Set .HorizontalRange = FXCell.Offset(3, 3).Resize(, ws.Range(FXCell.Offset(3, 3), ws.Cells(FXCell.Offset(3, 3).row, ws.Columns.Count)).End(xlToRight).Column - FXCell.Offset(3, 3).Column + 1)
        
        
    End With
    Dim DataString As String
    DataString = postCorrUpdater.CorrJsonh()
    Debug.Print DataString
    DataString = URLEncode(DataString)
    
    Dim url As String
    'url = "http://localhost:8080/val/marketdata/v1/corrs?baseDt=20231228&dataSetId=TEST11&matrixId=CORR"
    url = "http://localhost:8080/val/marketdata/v1/saveCorrs?baseDt=20231228&dataSetId=TEST16"
    ' JSON data와 POST request를 보내기 위해 subroutine을 호출한다.
    SendPostRequest DataString, url
        
        
        
End Sub

