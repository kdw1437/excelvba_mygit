Attribute VB_Name = "ClassPostYieldCurve"
'YieldCurve를 POST하는 코드(클래스 모듈 이용)
Sub ClassPostYieldCurve()

    Dim ycUpdater As PostYieldCurveUpdater
    Set ycUpdater = New PostYieldCurveUpdater

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    StartingPoint = Sheets("Market Data").Range("P2").Value 'StartingPoint 참조변수에 P2셀의 값 부여
        
    Dim Table1Point As Range
    Set Table1Point = Sheets("Market Data").Range(StartingPoint).Offset(3, 0) 'startingPoint(M4)명 (셀명)에서 3 row 밑의 셀을 Table1Point에 할당한다.
    
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, Table1Point.Column).End(xlUp).Row 'Table1Point와 같은 column에 있는 마지막으로 사용되어진 row를 찾는다.
    
    ' "Equity" table 다음에 "FX"를 포함하는 cell을 찾는다.
    Dim fxRow As Range 'Range객체는 하나 혹은 다수의 cell을 참조한다.
    Set fxRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(lastRow, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)
    'range안에 Table1Point 칼럼 영역 내에서 FX string을 찾는다. xlValues(formula가 아닌 cell값을 확인한다.) xlWhole (전체 셀의 내용이 찾는 단어와 완벽하게 일치함을 의미한다.)
    
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
    
    ' ArrayOfCurrency를 채운다.
    ycUpdater.PopulateArrayOfCurrency



    ' DataString을 만든다.
    Dim DataString As String
    DataString = ycUpdater.GenerateDataString

    
    Debug.Print DataString

    ' DataString을 encoding한다. (x-www-form-urlencoded)
    DataString = URLEncode(DataString)

    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/yieldcurves?baseDt=20231228&dataSetId=TEST13"

    ' POST request를 보낸다.
    SendPostRequest DataString, url

End Sub
