Attribute VB_Name = "ClassPostCorr"
Sub PrintJsonString()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")

    Dim equityCell As Range
    Set equityCell = ws.Columns("M:M").Find(What:="Equity", LookIn:=xlValues, LookAt:=xlWhole)

    ' postCorrUpdater�� ���ο� �ν��Ͻ� ����
    Dim postCorrUpdater As New postCorrUpdater
    With postCorrUpdater
        ' vertical range�� �����Ѵ�. (Equity cell�� ���� Ư�� ������ �ִ�)
        Set .VerticalRange = equityCell.Offset(4, 0).Resize(ws.Range(equityCell.Offset(4, 0), ws.Cells(ws.Rows.Count, equityCell.Column)).End(xlDown).Row - equityCell.Offset(4, 0).Row + 1)
        
        ' horizontal range�� �����Ѵ�. (Equity cell�� ���� Ư�� ������ �ִ�)
        Set .HorizontalRange = equityCell.Offset(3, 2).Resize(, ws.Range(equityCell.Offset(3, 2), ws.Cells(equityCell.Offset(3, 2).Row, ws.Columns.Count)).End(xlToRight).Column - equityCell.Offset(3, 2).Column + 1)
        
        
        'Debug.Print .CorrJson()
    End With
    Dim DataString As String
    DataString = postCorrUpdater.CorrJsonv()
    Debug.Print DataString
    DataString = URLEncode(DataString)
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/corrs?baseDt=20231228&dataSetId=TEST11&matrixId=CORR"
    
    ' JSON data�� POST request�� ������ ���� subroutine�� ȣ���Ѵ�.
    SendPostRequest DataString, url
End Sub

Sub PrintJsonString2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    
    Dim FXCell As Range
    Set FXCell = ws.Columns("M:M").Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)

    ' postCorrUpdater�� instance�� �����Ѵ�.
    Dim postCorrUpdater As New postCorrUpdater
    With postCorrUpdater
        ' VerticalRange property�� ���� �ο��Ѵ�. (setter�̿�)
        Set .VerticalRange = FXCell.Offset(4, 0).Resize(ws.Range(FXCell.Offset(4, 0), ws.Cells(ws.Rows.Count, FXCell.Column)).End(xlDown).Row - FXCell.Offset(4, 0).Row + 1)
        
        'HorizontalRange property�� ���� �ο��Ѵ�. (setter�̿�)
        Set .HorizontalRange = FXCell.Offset(3, 3).Resize(, ws.Range(FXCell.Offset(3, 3), ws.Cells(FXCell.Offset(3, 3).Row, ws.Columns.Count)).End(xlToRight).Column - FXCell.Offset(3, 3).Column + 1)
        
        
    End With
    Dim DataString As String
    DataString = postCorrUpdater.CorrJsonh()
    Debug.Print DataString
    DataString = URLEncode(DataString)
    
    
        
        
        
End Sub
