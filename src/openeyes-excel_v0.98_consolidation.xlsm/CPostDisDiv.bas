Attribute VB_Name = "CPostDisDiv"
Option Explicit
Dim requestHandler As CAsyncRequestHandler

Sub UseDividendDataProcessor()
    Dim dp As postDisDivMissingData
    Set dp = New postDisDivMissingData
    
    Set dp.Worksheet = ThisWorkbook.Worksheets("Missing Data - D_Dividend(보정)")
    Set dp.StartCell = dp.Worksheet.Range("A:A").Find(What:="Discrete Dividend", Lookat:=xlWhole)
    dp.k = 4 ' (K가 알려진 값이고 dynamic하게 결정되어질 때는 거기에 맞춰서 코드 작성)
    
    If Not dp.StartCell Is Nothing Then
        Dim jsonString As String
        
        jsonString = dp.ReturnJSON
        
        Debug.Print jsonString
    Else
        MsgBox "Start cell not found."
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Missing Data - D_Dividend(보정)")
    
    Dim baseDt As String
    baseDt = Format(ws.Range("B1").value, "yyyymmdd")
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveDividendStream?baseDt=" & baseDt & "&dataSetId=official"
    
    ' JSON data와 POST request를 보내는 subroutine을 호출한다.
    'SendPostRequest jsonString, url
    Set requestHandler = New CAsyncRequestHandler
    requestHandler.SendPostRequestAsync jsonString, url
End Sub


