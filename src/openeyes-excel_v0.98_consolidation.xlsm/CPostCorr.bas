Attribute VB_Name = "CPostCorr"
Option Explicit
Dim requestHandler As CAsyncRequestHandler

Sub UseCorrelationDataProcessor()
    Dim corrProcessor As New postCorrMissingData
    
    ' properties 세팅
    Set corrProcessor.Worksheet = ThisWorkbook.Sheets("Missing Data - Hist Vol, Corr")
    corrProcessor.startRow = 5
    corrProcessor.Column = "E"
    
    ' Generate JSON
    Dim jsonString As String
    jsonString = corrProcessor.GenerateJSON
    
    ' Print JSON
    Debug.Print jsonString
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Missing Data - Hist Vol, Corr")
    
    Dim dataSetId As String
    dataSetId = "official"
    
    Dim baseDt As String
    baseDt = Format(ws.Range("B1").value, "yyyymmdd")
    
    Dim url As String
    
    url = "http://localhost:8080/val/marketdata/v1/saveCorrs?baseDt=" & baseDt & "&dataSetId=" & dataSetId & "&matrixId=CORR"
    ' JSON data와 POST request를 보내기 위해 subroutine을 호출한다.
    'SendPostRequest DataString, url
    Set requestHandler = New CAsyncRequestHandler
    ' POST request를 보낸다.
    'SendPostRequest DataString, url
    requestHandler.SendPostRequestAsync jsonString, url
    
End Sub
