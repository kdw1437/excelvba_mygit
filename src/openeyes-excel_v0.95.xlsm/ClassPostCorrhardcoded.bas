Attribute VB_Name = "ClassPostCorrhardcoded"
Sub PrintJsonString()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    
    Dim equityCell As Range
    Set equityCell = ws.Range(ws.Range("P2").value)
    ' PostCorrUpdater class의 인스턴스 생성
    Dim postCorrUpdater As New postCorrUpdater
    With postCorrUpdater
        ' vertical ranget와 horizontal range setter 호출
        Set .VerticalRange = ws.Range(equityCell.Offset(4, 0), equityCell.Offset(4, 0).End(xlDown))
        Set .HorizontalRange = ws.Range(equityCell.Offset(3, 2), equityCell.Offset(3, 2).End(xlToRight))
        
'        Set .VerticalRange = ws.Range("M8", ws.Range("M8").End(xlDown))
'        Set .HorizontalRange = ws.Range("O7", ws.Range("O7").End(xlToRight))
        
        
        'Debug.Print .CorrJson()
    End With
    Dim DataString As String
    DataString = postCorrUpdater.CorrJsonv2()
    Debug.Print DataString
    DataString = URLEncode(DataString)
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/corrs?baseDt=20231228&dataSetId=TEST11&matrixId=CORR"
    
    ' JSON data와 POST request를 보내기 위해 subroutine을 호출한다.
    SendPostRequest DataString, url
End Sub

Sub PrintJsonString2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    Dim FXCellColumn As Long
    FXCellColumn = ws.Range(ws.Range("P2").value).Column
    
    Dim FXCell As Range
    Dim searchRange As Range
    Set searchRange = ws.Columns(FXCellColumn)
    
    Set FXCell = searchRange.Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    
    Dim postCorrUpdater As New postCorrUpdater
    With postCorrUpdater
    
        Set .VerticalRange = ws.Range(FXCell.Offset(4, 0), FXCell.Offset(4, 0).End(xlDown))
        Set .HorizontalRange = ws.Range(FXCell.Offset(3, 3), FXCell.Offset(3, 3).End(xlToRight))
        
    End With
    Dim DataString As String
    DataString = postCorrUpdater.CorrJsonh2()
    Debug.Print DataString
    DataString = URLEncode(DataString)
    
    
        
        
        
End Sub
