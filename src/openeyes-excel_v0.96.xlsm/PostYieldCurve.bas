Attribute VB_Name = "PostYieldCurve"
'YieldCurve�� POST�ϴ� �ڵ�
Sub PostYieldCurve()
    
    Dim i As Integer
    'Dim baseDt As String
    Dim dataSetId As String
    Dim StartingPoint As String
    Dim dataId As String
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data") 'ws���� ��, ws ���������� Market Data sheet�� ����Ŵ.
    'Dim targetDate As Date
    ' Retrieve the base date and data set ID from the worksheet
    'targetDate = Sheets("Market Data").Range("A2").Value
    
    'baseDt = Format(targetDate, "yyyymmdd")
    dataSetId = Sheets("Market Data").Range("O2").value 'dataSetId ���������� O2���� �� �ο�
    StartingPoint = Sheets("Market Data").Range("P2").value 'StartingPoint ���������� P2���� �� �ο�
        
    Dim Table1Point As Range
    Set Table1Point = Sheets("Market Data").Range(StartingPoint).Offset(3, 0) 'startingPoint(M4)�� (����)���� 3 row ���� ���� Table1Point�� �Ҵ��Ѵ�.
    
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, Table1Point.Column).End(xlUp).row 'Table1Point�� ���� column�� �ִ� ���������� ���Ǿ��� row�� ã�´�.
    
    ' Find the cell that contains "FX" after "Equity" table
    Dim fxRow As Range 'Range��ü�� �ϳ� Ȥ�� �ټ��� cell�� �����Ѵ�.
    Set fxRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(lastRow, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)
    'range�ȿ� Table1Point Į�� ���� ������ FX string�� ã�´�. xlValues(formula�� �ƴ� cell���� Ȯ���Ѵ�.) xlWhole (��ü ���� ������ ã�� �ܾ�� �Ϻ��ϰ� ��ġ���� �ǹ��Ѵ�.)
    
    Dim Table2Point As Range
    Set Table2Point = fxRow.Offset(3, 0)
    
    Dim YieldCurveRow As Range
    Set YieldCurveRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(lastRow, Table1Point.Column)).Find(What:="Yield Curve", LookIn:=xlValues, LookAt:=xlWhole)
    'Debug.Print Table2Point.value

    Dim DATA_ID_Cell1 As Range
    Set DATA_ID_Cell1 = ws.Cells(YieldCurveRow.row + 2, YieldCurveRow.Column)
    'Debug.Print DATA_ID_Cell1.value
    Dim DATA_ID_Cells() As Variant
    Dim colIndex As Long
    Dim currentCell As Range
    Dim cellCount As Integer
    
    Set currentCell = DATA_ID_Cell1
    cellCount = 0
    
   '�� ���� ���� ������ loop�� ���ư���.
    Do
        '���� ���� ������� �ƴ��� Ȯ���Ѵ�.
        If IsEmpty(currentCell.value) Then
            Exit Do '�� ���� �߰� �Ǿ��� ��, loop���� �������´�.
        End If
        
        ' array�� resize�ϰ�, ���� ���� ���� array�� �Ҵ��Ѵ�.
        cellCount = cellCount + 1
        ReDim Preserve DATA_ID_Cells(1 To cellCount)
        DATA_ID_Cells(cellCount) = currentCell.value
        
        ' Move to the next cell 2 columns to the right
        Set currentCell = ws.Cells(currentCell.row, currentCell.Column + 2)
    Loop
    Dim arraySize As Integer
    arraySize = UBound(DATA_ID_Cells) 'arraySize�� DATA_ID_Cells�� size�� �Ҵ��Ѵ�.
    Dim InterestName As String
    Dim j As Integer
    Dim tenor As Double
    Dim Rate As Double
    Dim RiskCode As String
    Dim DataString As String
    ' DataString�� �ʱ�ȭ�Ѵ�.
    DataString = "["
    If arraySize > 0 Then
        For i = 1 To arraySize
            InterestName = DATA_ID_Cells(i) 'For���� arraySize��ŭ ������. DATA_ID_Cells�� �� ���Ҹ� InterestName�� �Ҵ��Ѵ�.
            j = 1 'inner loop�� counter�� �ʱ�ȭ�Ѵ�.
    
            ' Yields array�� �����.
            Dim yieldsArray As String
            yieldsArray = "["
            Do While Not IsEmpty(ws.Cells(YieldCurveRow.row + 3 + j, YieldCurveRow.Column + (i - 1) * 2))
                tenor = ws.Cells(YieldCurveRow.row + 3 + j, YieldCurveRow.Column + (i - 1) * 2).value 'value���� Ư�� ���� (Double type)�� �Ҵ��Ѵ�.
                Rate = ws.Cells(YieldCurveRow.row + 3 + j, YieldCurveRow.Column + (i - 1) * 2 + 1).value
                RiskCode = Format(tenor * 360, "00000")
    
                ' yield ��ü�� ����� �̸� array yield�� �Ҵ��Ѵ�.
                If yieldsArray <> "[" Then yieldsArray = yieldsArray & ","
                yieldsArray = yieldsArray & "{""tenor"": " & tenor & ", ""rate"": " & Rate & "}" 'Ư�� �������� json string���� �����Ѵ�.
    
                j = j + 1
            Loop
            yieldsArray = yieldsArray & "]"
    
            ' main JSON object�� InterestName�� Yields array�� �߰��Ѵ�.
            If i > 1 Then DataString = DataString & ","
            DataString = DataString & "{" & _
                                     """dataId"": """ & InterestName & """, " & _
                                     """currency"": """ & Left(InterestName, 3) & """, " & _
                                     """yields"": " & yieldsArray & "}"
        Next i
    End If
    DataString = DataString & "]"

    
        
    Debug.Print DataString

    ' (x-www-form-urlencoded) DataString�� Encoding�Ѵ�.
    DataString = URLEncode(DataString)
    
    
    ' request�� ���� URL
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/yieldcurves?baseDt=20231228&dataSetId=TEST11"
    
    SendPostRequest DataString, url

End Sub




