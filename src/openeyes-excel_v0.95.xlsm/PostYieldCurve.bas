Attribute VB_Name = "PostYieldCurve"
'YieldCurve를 POST하는 코드
Sub PostYieldCurve()
    
    Dim i As Integer
    'Dim baseDt As String
    Dim dataSetId As String
    Dim StartingPoint As String
    Dim dataId As String
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data") 'ws선언 후, ws 참조변수로 Market Data sheet을 가리킴.
    'Dim targetDate As Date
    ' Retrieve the base date and data set ID from the worksheet
    'targetDate = Sheets("Market Data").Range("A2").Value
    
    'baseDt = Format(targetDate, "yyyymmdd")
    dataSetId = Sheets("Market Data").Range("O2").value 'dataSetId 참조변수에 O2셀의 값 부여
    StartingPoint = Sheets("Market Data").Range("P2").value 'StartingPoint 참조변수에 P2셀의 값 부여
        
    Dim Table1Point As Range
    Set Table1Point = Sheets("Market Data").Range(StartingPoint).Offset(3, 0) 'startingPoint(M4)명 (셀명)에서 3 row 밑의 셀을 Table1Point에 할당한다.
    
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, Table1Point.Column).End(xlUp).row 'Table1Point와 같은 column에 있는 마지막으로 사용되어진 row를 찾는다.
    
    ' Find the cell that contains "FX" after "Equity" table
    Dim fxRow As Range 'Range객체는 하나 혹은 다수의 cell을 참조한다.
    Set fxRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(lastRow, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)
    'range안에 Table1Point 칼럼 영역 내에서 FX string을 찾는다. xlValues(formula가 아닌 cell값을 확인한다.) xlWhole (전체 셀의 내용이 찾는 단어와 완벽하게 일치함을 의미한다.)
    
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
    
   '빈 셀이 나올 때까지 loop가 돌아간다.
    Do
        '현재 셀이 비었는지 아닌지 확인한다.
        If IsEmpty(currentCell.value) Then
            Exit Do '빈 셀이 발견 되었을 때, loop에서 빠져나온다.
        End If
        
        ' array를 resize하고, 현재 셀의 값을 array에 할당한다.
        cellCount = cellCount + 1
        ReDim Preserve DATA_ID_Cells(1 To cellCount)
        DATA_ID_Cells(cellCount) = currentCell.value
        
        ' Move to the next cell 2 columns to the right
        Set currentCell = ws.Cells(currentCell.row, currentCell.Column + 2)
    Loop
    Dim arraySize As Integer
    arraySize = UBound(DATA_ID_Cells) 'arraySize에 DATA_ID_Cells의 size를 할당한다.
    Dim InterestName As String
    Dim j As Integer
    Dim tenor As Double
    Dim Rate As Double
    Dim RiskCode As String
    Dim DataString As String
    ' DataString을 초기화한다.
    DataString = "["
    If arraySize > 0 Then
        For i = 1 To arraySize
            InterestName = DATA_ID_Cells(i) 'For문을 arraySize만큼 돌린다. DATA_ID_Cells의 각 원소를 InterestName에 할당한다.
            j = 1 'inner loop의 counter를 초기화한다.
    
            ' Yields array를 만든다.
            Dim yieldsArray As String
            yieldsArray = "["
            Do While Not IsEmpty(ws.Cells(YieldCurveRow.row + 3 + j, YieldCurveRow.Column + (i - 1) * 2))
                tenor = ws.Cells(YieldCurveRow.row + 3 + j, YieldCurveRow.Column + (i - 1) * 2).value 'value값을 특정 변수 (Double type)에 할당한다.
                Rate = ws.Cells(YieldCurveRow.row + 3 + j, YieldCurveRow.Column + (i - 1) * 2 + 1).value
                RiskCode = Format(tenor * 360, "00000")
    
                ' yield 객체를 만들고 이를 array yield에 할당한다.
                If yieldsArray <> "[" Then yieldsArray = yieldsArray & ","
                yieldsArray = yieldsArray & "{""tenor"": " & tenor & ", ""rate"": " & Rate & "}" '특정 변수값을 json string으로 연결한다.
    
                j = j + 1
            Loop
            yieldsArray = yieldsArray & "]"
    
            ' main JSON object에 InterestName과 Yields array를 추가한다.
            If i > 1 Then DataString = DataString & ","
            DataString = DataString & "{" & _
                                     """dataId"": """ & InterestName & """, " & _
                                     """currency"": """ & Left(InterestName, 3) & """, " & _
                                     """yields"": " & yieldsArray & "}"
        Next i
    End If
    DataString = DataString & "]"

    
        
    Debug.Print DataString

    ' (x-www-form-urlencoded) DataString을 Encoding한다.
    DataString = URLEncode(DataString)
    
    
    ' request를 보낼 URL
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/yieldcurves?baseDt=20231228&dataSetId=TEST11"
    
    SendPostRequest DataString, url

End Sub




