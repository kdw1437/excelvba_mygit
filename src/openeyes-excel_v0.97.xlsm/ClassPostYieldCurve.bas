Attribute VB_Name = "ClassPostYieldCurve"
Dim requestHandler As CAsyncRequestHandler
'YieldCurve�� POST�ϴ� �ڵ�(Ŭ���� ��� �̿�)
Sub ClassPostYieldCurve()

    Dim ycUpdater As PostYieldCurveUpdater
    Set ycUpdater = New PostYieldCurveUpdater

    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("Market Data")
    StartingPoint = Sheets("Market Data").Range("P2").value 'StartingPoint ���������� P2���� �� �ο�
        
    Dim Table1Point As Range
    Set Table1Point = Sheets("Market Data").Range(StartingPoint).Offset(3, 0) 'startingPoint(M4)�� (����)���� 3 row ���� ���� Table1Point�� �Ҵ��Ѵ�.
    
    Dim lastRow As Long
    
    lastRow = Ws.Cells(Ws.Rows.Count, Table1Point.Column).End(xlUp).row 'Table1Point�� ���� column�� �ִ� ���������� ���Ǿ��� row�� ã�´�.
    
    ' "Equity" table ������ "FX"�� �����ϴ� cell�� ã�´�.
    Dim fxRow As Range 'Range��ü�� �ϳ� Ȥ�� �ټ��� cell�� �����Ѵ�.
    Set fxRow = Ws.Range(Table1Point.Offset(1, 0), Ws.Cells(lastRow, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)
    'range�ȿ� Table1Point Į�� ���� ������ FX string�� ã�´�. xlValues(formula�� �ƴ� cell���� Ȯ���Ѵ�.) xlWhole (��ü ���� ������ ã�� �ܾ�� �Ϻ��ϰ� ��ġ���� �ǹ��Ѵ�.)
    
    Dim Table2Point As Range
    Set Table2Point = fxRow.Offset(3, 0)
    
    Dim YieldCurveRow As Range
    Set YieldCurveRow = Ws.Range(Table1Point.Offset(1, 0), Ws.Cells(lastRow, Table1Point.Column)).Find(What:="Yield Curve", LookIn:=xlValues, LookAt:=xlWhole)
    'Debug.Print Table2Point.value

    Dim DATA_ID_Cell1 As Range
    Set DATA_ID_Cell1 = Ws.Cells(YieldCurveRow.row + 2, YieldCurveRow.Column)
    ' Set StartCell
    
    
    Set ycUpdater.startCell = DATA_ID_Cell1


    ' Set YieldCurveRow

    Set ycUpdater.YieldCurveRow = YieldCurveRow
    
    ' ArrayOfCurrency�� ä���.
    ycUpdater.PopulateArrayOfCurrency



    ' DataString�� �����.
    Dim DataString As String
    DataString = ycUpdater.GenerateDataString2

    
    Debug.Print DataString

    ' DataString�� encoding�Ѵ�. (x-www-form-urlencoded)
    DataString = URLEncode(DataString)

    Dim dataSetId As String
    dataSetId = Ws.Range("O2").value
    
    Dim baseDt As String
    baseDt = Format(Ws.Range("A2").value, "yyyymmdd")
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveYieldcurves?baseDt=" & baseDt & "&dataSetId=" & dataSetId
    'url = "http://localhost:8080/val/marketdata/v1/saveYieldcurves?baseDt=20231229"
    
    Set requestHandler = New CAsyncRequestHandler
    ' POST request�� ������.
    'SendPostRequest DataString, url
    requestHandler.SendPostRequestAsync DataString, url
End Sub