Attribute VB_Name = "ClassPostCorr"
Sub PrintJsonString()
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("Market Data")

    Dim equityCell As Range
    Set equityCell = Ws.Columns("M:M").Find(What:="Equity", LookIn:=xlValues, LookAt:=xlWhole)

    ' postCorrUpdater�� ���ο� �ν��Ͻ� ����
    Dim postCorrUpdater As New postCorrUpdater
    With postCorrUpdater
        ' vertical range�� �����Ѵ�. (Equity cell�� ���� Ư�� ������ �ִ�)
        Set .VerticalRange = equityCell.Offset(4, 0).Resize(Ws.Range(equityCell.Offset(4, 0), Ws.Cells(Ws.Rows.Count, equityCell.Column)).End(xlDown).row - equityCell.Offset(4, 0).row + 1)
        
        ' horizontal range�� �����Ѵ�. (Equity cell�� ���� Ư�� ������ �ִ�)
        Set .HorizontalRange = equityCell.Offset(3, 2).Resize(, Ws.Range(equityCell.Offset(3, 2), Ws.Cells(equityCell.Offset(3, 2).row, Ws.Columns.Count)).End(xlToRight).Column - equityCell.Offset(3, 2).Column + 1)
        
        
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
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("Market Data")
    
    Dim FXCell As Range
    Set FXCell = Ws.Columns("M:M").Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)

    ' postCorrUpdater�� instance�� �����Ѵ�.
    Dim postCorrUpdater As New postCorrUpdater
    With postCorrUpdater
        ' VerticalRange property�� ���� �ο��Ѵ�. (setter�̿�)
        Set .VerticalRange = FXCell.Offset(4, 0).Resize(Ws.Range(FXCell.Offset(4, 0), Ws.Cells(Ws.Rows.Count, FXCell.Column)).End(xlDown).row - FXCell.Offset(4, 0).row + 1)
        
        'HorizontalRange property�� ���� �ο��Ѵ�. (setter�̿�)
        Set .HorizontalRange = FXCell.Offset(3, 3).Resize(, Ws.Range(FXCell.Offset(3, 3), Ws.Cells(FXCell.Offset(3, 3).row, Ws.Columns.Count)).End(xlToRight).Column - FXCell.Offset(3, 3).Column + 1)
        
        
    End With
    Dim DataString As String
    DataString = postCorrUpdater.CorrJsonh()
    Debug.Print DataString
    DataString = URLEncode(DataString)
    
    
        
        
        
End Sub

