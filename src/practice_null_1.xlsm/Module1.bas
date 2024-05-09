Attribute VB_Name = "Module1"
Option Explicit

Sub Practice()


    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    Dim i As Integer
    Dim ToJsonData As Collection
    Set ToJsonData = New Collection
    
    For i = 1 To 4
        Dim ToJsonDataDic As Dictionary
        Set ToJsonDataDic = New Dictionary
        ToJsonDataDic.Add "Tenor", IfEmpty(CStr(ws.Cells(i + 1, 1).value))
        ToJsonDataDic.Add "Rate", IfEmpty(CStr(ws.Cells(i + 1, 2).value))
        ToJsonData.Add ToJsonDataDic
    Next i
    
    Dim JsonString As String
    JsonString = JsonConverter.ConvertToJson(ToJsonData)
        
    Debug.Print JsonString

End Sub

' Helper function to convert empty cells to Null
Function IfEmpty(value As Variant) As Variant
    If IsEmpty(value) Or value = "" Then
        IfEmpty = Null  ' Assign Null if cell is empty
    Else
        IfEmpty = value
    End If
End Function
