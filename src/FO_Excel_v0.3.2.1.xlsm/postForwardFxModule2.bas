Attribute VB_Name = "postForwardFxModule2"
Option Explicit

Sub postForwardFXModule()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Missing Data - Fx Forward")
    
    Dim StartCell As Range
    Set StartCell = ws.Range("A:A").Find(What:="FX Forward Curve", Lookat:=xlWhole)
    
    Dim i As Long
    Dim relCurrencyCol As Collection
    Set relCurrencyCol = New Collection
    Dim k As Long
    k = 4
    
    For i = 1 To k '이거 column dynamic하게 들어오는 것을 대비해서 To 다음 4가 수정되야 함.
        relCurrencyCol.Add StartCell.Offset(4, 1 + 3 * (i - 1)).value
    Next i
    
    Dim j As Long
    Dim CurrencyCol As Collection
    Set CurrencyCol = New Collection
    
    For j = 1 To k
        CurrencyCol.Add StartCell.Offset(3, 1 + 3 * (j - 1)).value
    Next j
    
    Dim jsonString As String
    jsonString = "["
    
    For i = 1 To k
        Dim relCurrencyCell As Range
        Set relCurrencyCell = ws.Range("7:7").Find(What:=relCurrencyCol(i), Lookat:=xlWhole)
        jsonString = jsonString & "{" & Chr(34) & "dataId" & Chr(34) & ": " & Chr(34) & "F_FX_" & relCurrencyCol(i) & CurrencyCol(i) & Chr(34) & ", " & Chr(34) & "yields" & Chr(34) & ": ["
        j = 2
        Do While ws.Cells(relCurrencyCell.Row + j, relCurrencyCell.Column).value <> ""
            Dim tenor As Double
            Dim value As Double
            tenor = ws.Cells(relCurrencyCell.Row + j, relCurrencyCell.Column - 1).value
            value = ws.Cells(relCurrencyCell.Row + j, relCurrencyCell.Column).value
            
            ' Append the tenor-value pair to the JSON object
            jsonString = jsonString & "{" & Chr(34) & "tenor" & Chr(34) & ": " & tenor & ", " & Chr(34) & "value" & Chr(34) & ": " & value & "},"
            
            '[{"dataId": F_FX + relCurrencyCell.value + relCurrencyCell.offset(-1,0).value, "yields" : [{"tenor": 0.00278, "value": 1},{"tenor": 0.25, "value: 2}....]
            j = j + 1
        Loop
        
        jsonString = Left(jsonString, Len(jsonString) - 1) & "]}"
        
        ' Add a comma between JSON objects if not at the last object
        If i < k Then
            jsonString = jsonString & ","
        End If
    Next i
    jsonString = jsonString & "]"
    Debug.Print jsonString
    
End Sub


Sub postForwardFXModule3()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Missing Data - Fx Forward")
    
    Dim StartCell As Range
    Set StartCell = ws.Range("A:A").Find(What:="FX Forward Curve", Lookat:=xlWhole)
    
    Dim currenciesCollection As New Collection
    Dim i As Long, j As Long, k As Long
    k = 4  ' This should be dynamically determined if possible
    
    ' Loop through each currency block and collect the data
    For i = 1 To k
        Dim relCurrencyCell As Range
        Dim currencyCode As String
        currencyCode = StartCell.Offset(4, 1 + 3 * (i - 1)).value
        
        Set relCurrencyCell = ws.Range("7:7").Find(What:=currencyCode, Lookat:=xlWhole)
        
        Dim currencyData As Object
        Set currencyData = CreateObject("Scripting.Dictionary")
        Dim dataId As String
        dataId = "F_FX_" & currencyCode & StartCell.Offset(3, 1 + 3 * (i - 1)).value
        currencyData("dataId") = dataId
        
        Dim yieldsCollection As New Collection
        Set yieldsCollection = New Collection 'Reinitialize
        
        j = 2 ' Start row offset for yields
        Do While ws.Cells(relCurrencyCell.Row + j, relCurrencyCell.Column).value <> ""
            Dim yieldData As Object
            Set yieldData = CreateObject("Scripting.Dictionary")
            
            yieldData("tenor") = ws.Cells(relCurrencyCell.Row + j, relCurrencyCell.Column - 1).value
            yieldData("value") = ws.Cells(relCurrencyCell.Row + j, relCurrencyCell.Column).value
            
            yieldsCollection.Add yieldData
            j = j + 1
        Loop
        'Collection은 dictionary의 key의 value가 될 수 없다.
        'collection을 array로 바꿔 주고, array를 value로 key와 함께 할당한다.
        Dim yieldsArray() As Variant
        ReDim yieldsArray(1 To yieldsCollection.Count)
        
        Dim idx As Long
        For idx = 1 To yieldsCollection.Count
            Set yieldsArray(idx) = yieldsCollection(idx)
        Next idx
        
        '이제 array를 dictionary에 할당한다.
        currencyData("yields") = yieldsArray
        currenciesCollection.Add currencyData
    Next i
    
    ' 전체 collection을 JSON으로 변경한다.
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(currenciesCollection)
    
    Debug.Print jsonString
End Sub

Sub UseFXForwardData()
    Dim fxData As New PostForwardRate
    
    Set fxData.Worksheet = ThisWorkbook.Sheets("Missing Data - Fx Forward")
    Set fxData.StartCell = fxData.Worksheet.Range("A:A").Find(What:="FX Forward Curve", Lookat:=xlWhole)
    fxData.k = 4  ' Setting how many currencies to process
    
    'fxData.GenerateJSON
    
    Dim jsonString As String
    jsonString = fxData.ReturnJSON
    'Debug.Print jsonString
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveForwardFX?baseDt=20240412&dataSetId=TEST2"
    
    ' JSON data와 POST request를 보내는 subroutine을 호출한다.
    SendPostRequest jsonString, url
End Sub

