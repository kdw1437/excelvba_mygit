Attribute VB_Name = "PostPrice"
Sub PostPrice()
    Dim DataString As String
    Dim i As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    
    Dim StartingPoint As Range
    Set StartingPoint = ws.Range(ws.Range("P2").Value)
    
    Dim Table1Point As Range
    Set Table1Point = StartingPoint.Offset(3, 0)
    
    ' "Equity" table 다음에 "FX"를 포함하는 cell을 찾는다.
    Dim fxRow As Range
    Set fxRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(ws.Rows.Count, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)

    ' 빈 JSON array로 DataString을 초기화한다.
    DataString = "["

    For i = Table1Point.Row + 1 To fxRow.Row - 2
        Dim DataId As String
        Dim closePric As String
        
        DataId = ws.Cells(i, Table1Point.Column).Value
        closePric = ws.Cells(i, Table1Point.Column + 1).Value
        
        ' 현재 row에 대해 JSON object를 만든다.
        Dim jsonObject As String
        jsonObject = "{""dataId"": """ & DataId & """, ""price"": " & closePric & "}"
        
        ' 첫번째 item이 아니라면, comma 분리자를 추가한다.
        If Len(DataString) > 1 Then
            DataString = DataString & ", "
        End If
        
        ' (JSON array) JSON 객체를 DataString(JSON array)에 붙인다.
        DataString = DataString & jsonObject
    Next i

    ' JSON array를 닫는다.
    DataString = DataString & "]"
    
    Debug.Print DataString
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/prices?baseDt=20231228&dataSetId=TEST11"
    
    ' JSON data와 POST request를 보내는 subroutine을 호출한다.
    SendPostRequest DataString, url
End Sub


