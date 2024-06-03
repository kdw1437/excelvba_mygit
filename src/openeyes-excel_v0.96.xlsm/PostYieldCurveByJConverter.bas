Attribute VB_Name = "PostYieldCurveByJConverter"
Sub PostYieldCurve()
    
    Dim i As Integer
    Dim dataSetId As String
    Dim StartingPoint As String
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    
    dataSetId = Sheets("Market Data").Range("O2").value
    StartingPoint = Sheets("Market Data").Range("P2").value
        
    Dim Table1Point As Range
    Set Table1Point = Sheets("Market Data").Range(StartingPoint).Offset(3, 0)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, Table1Point.Column).End(xlUp).row
    
    Dim fxRow As Range
    Set fxRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(lastRow, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)
    
    Dim Table2Point As Range
    Set Table2Point = fxRow.Offset(3, 0)
    
    Dim YieldCurveRow As Range
    Set YieldCurveRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(lastRow, Table1Point.Column)).Find(What:="Yield Curve", LookIn:=xlValues, LookAt:=xlWhole)
    
    Dim DATA_ID_Cell1 As Range
    Set DATA_ID_Cell1 = ws.Cells(YieldCurveRow.row + 2, YieldCurveRow.Column)
    
    Dim DATA_ID_Cells() As Variant
    Dim colIndex As Long
    Dim currentCell As Range
    Dim cellCount As Integer
    
    Set currentCell = DATA_ID_Cell1
    cellCount = 0
    
    Do
        If IsEmpty(currentCell.value) Then
            Exit Do
        End If
        
        cellCount = cellCount + 1
        ReDim Preserve DATA_ID_Cells(1 To cellCount)
        DATA_ID_Cells(cellCount) = currentCell.value
        
        Set currentCell = ws.Cells(currentCell.row, currentCell.Column + 2)
    Loop
    
    Dim arraySize As Integer
    arraySize = UBound(DATA_ID_Cells)
    
    Dim InterestName As String
    Dim j As Integer
    Dim tenor As Double
    Dim Rate As Double
    Dim RiskCode As String
    
    Dim jsonArray As New Collection
    
    If arraySize > 0 Then
        For i = 1 To arraySize
            InterestName = DATA_ID_Cells(i)
            j = 1
            
            Dim yieldsArray As New Collection
            Set yieldsArray = New Collection
            Do While Not IsEmpty(ws.Cells(YieldCurveRow.row + 3 + j, YieldCurveRow.Column + (i - 1) * 2))
                Dim yieldObj As New Dictionary
                Set yieldObj = New Dictionary
                
                tenor = ws.Cells(YieldCurveRow.row + 3 + j, YieldCurveRow.Column + (i - 1) * 2).value
                Rate = ws.Cells(YieldCurveRow.row + 3 + j, YieldCurveRow.Column + (i - 1) * 2 + 1).value
                RiskCode = Format(tenor * 360, "00000")
                
                ' Create a new Dictionary for each yield object with both "tenor" and "rate" keys
                
                yieldObj.Add "tenor", tenor
                yieldObj.Add "rate", Rate
                
                yieldsArray.Add yieldObj
                
                j = j + 1
            Loop
            
            Dim dataObj As New Dictionary
            Set dataObj = New Dictionary
            
            dataObj.Add "dataId", InterestName
            dataObj.Add "currency", Left(InterestName, 3)
            dataObj.Add "yields", yieldsArray
            
            jsonArray.Add dataObj
        Next i
    End If
    
    Dim JsonString As String
    JsonString = JsonConverter.ConvertToJson(jsonArray)
    
    Debug.Print JsonString
    
    Dim EncodedString As String
    EncodedString = URLEncode(JsonString)
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/yieldcurves?baseDt=20231228&dataSetId=TEST11"
    
    SendPostRequest EncodedString, url

End Sub
