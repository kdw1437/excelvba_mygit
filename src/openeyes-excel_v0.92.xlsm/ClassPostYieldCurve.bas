Attribute VB_Name = "ClassPostYieldCurve"
'YieldCurve�� POST�ϴ� �ڵ�(Ŭ���� ��� �̿�)
Sub ClassPostYieldCurve()

    Dim ycUpdater As PostYieldCurveUpdater
    Set ycUpdater = New PostYieldCurveUpdater

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    StartingPoint = Sheets("Market Data").Range("P2").Value 'StartingPoint ���������� P2���� �� �ο�
        
    Dim Table1Point As Range
    Set Table1Point = Sheets("Market Data").Range(StartingPoint).Offset(3, 0) 'startingPoint(M4)�� (����)���� 3 row ���� ���� Table1Point�� �Ҵ��Ѵ�.
    
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, Table1Point.Column).End(xlUp).Row 'Table1Point�� ���� column�� �ִ� ���������� ���Ǿ��� row�� ã�´�.
    
    ' "Equity" table ������ "FX"�� �����ϴ� cell�� ã�´�.
    Dim fxRow As Range 'Range��ü�� �ϳ� Ȥ�� �ټ��� cell�� �����Ѵ�.
    Set fxRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(lastRow, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)
    'range�ȿ� Table1Point Į�� ���� ������ FX string�� ã�´�. xlValues(formula�� �ƴ� cell���� Ȯ���Ѵ�.) xlWhole (��ü ���� ������ ã�� �ܾ�� �Ϻ��ϰ� ��ġ���� �ǹ��Ѵ�.)
    
    Dim Table2Point As Range
    Set Table2Point = fxRow.Offset(3, 0)
    
    Dim YieldCurveRow As Range
    Set YieldCurveRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(lastRow, Table1Point.Column)).Find(What:="Yield Curve", LookIn:=xlValues, LookAt:=xlWhole)
    'Debug.Print Table2Point.value

    Dim DATA_ID_Cell1 As Range
    Set DATA_ID_Cell1 = ws.Cells(YieldCurveRow.Row + 2, YieldCurveRow.Column)
    ' Set StartCell
    
    
    Set ycUpdater.startCell = DATA_ID_Cell1


    ' Set YieldCurveRow

    Set ycUpdater.YieldCurveRow = YieldCurveRow
    
    ' ArrayOfCurrency�� ä���.
    ycUpdater.PopulateArrayOfCurrency



    ' DataString�� �����.
    Dim DataString As String
    DataString = ycUpdater.GenerateDataString

    
    Debug.Print DataString

    ' DataString�� encoding�Ѵ�. (x-www-form-urlencoded)
    DataString = URLEncode(DataString)

    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/yieldcurves?baseDt=20231228&dataSetId=TEST13"

    ' POST request�� ������.
    SendPostRequest DataString, url

End Sub
