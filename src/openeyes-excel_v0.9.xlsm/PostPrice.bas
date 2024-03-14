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
    
    ' "Equity" table ������ "FX"�� �����ϴ� cell�� ã�´�.
    Dim fxRow As Range
    Set fxRow = ws.Range(Table1Point.Offset(1, 0), ws.Cells(ws.Rows.Count, Table1Point.Column)).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)

    ' �� JSON array�� DataString�� �ʱ�ȭ�Ѵ�.
    DataString = "["

    For i = Table1Point.Row + 1 To fxRow.Row - 2
        Dim DataId As String
        Dim closePric As String
        
        DataId = ws.Cells(i, Table1Point.Column).Value
        closePric = ws.Cells(i, Table1Point.Column + 1).Value
        
        ' ���� row�� ���� JSON object�� �����.
        Dim jsonObject As String
        jsonObject = "{""dataId"": """ & DataId & """, ""price"": " & closePric & "}"
        
        ' ù��° item�� �ƴ϶��, comma �и��ڸ� �߰��Ѵ�.
        If Len(DataString) > 1 Then
            DataString = DataString & ", "
        End If
        
        ' (JSON array) JSON ��ü�� DataString(JSON array)�� ���δ�.
        DataString = DataString & jsonObject
    Next i

    ' JSON array�� �ݴ´�.
    DataString = DataString & "]"
    
    Debug.Print DataString
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/prices?baseDt=20231228&dataSetId=TEST11"
    
    ' JSON data�� POST request�� ������ subroutine�� ȣ���Ѵ�.
    SendPostRequest DataString, url
End Sub


