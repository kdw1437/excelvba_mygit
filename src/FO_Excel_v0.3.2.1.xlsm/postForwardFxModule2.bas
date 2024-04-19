Attribute VB_Name = "postForwardFxModule2"
Option Explicit

Sub postForwardFXModule()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Missing Data - Fx Forward")
    
    Dim startCell As Range
    Set startCell = ws.Range("A:A").Find(What:="FX Forward Curve", Lookat:=xlWhole)
    
    Dim i As Long
    Dim relCurrencyCol As Collection
    Set relCurrencyCol = New Collection
    Dim k As Long
    k = 4
    
    For i = 1 To k '이거 column dynamic하게 들어오는 것을 대비해서 To 다음 4가 수정되야 함.
        relCurrencyCol.Add startCell.Offset(4, 1 + 3 * (i - 1)).value
    Next i
    
    Dim j As Long
    Dim CurrencyCol As Collection
    Set CurrencyCol = New Collection
    
    For j = 1 To k
        CurrencyCol.Add startCell.Offset(3, 1 + 3 * (j - 1)).value
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

