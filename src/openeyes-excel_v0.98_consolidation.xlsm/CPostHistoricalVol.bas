Attribute VB_Name = "CPostHistoricalVol"
Option Explicit
Dim requestHandler As CAsyncRequestHandler

Sub UseHistoricalVolProcessor()
    Dim volProcessor As postHistoricalVol
    Set volProcessor = New postHistoricalVol
    
    ' Set properties
    Set volProcessor.Worksheet = ThisWorkbook.Sheets("Missing Data - Hist Vol, Corr")
    volProcessor.startRow = 5
    
    ' Call the method to process data
    Dim jsonString As String
    jsonString = volProcessor.ReturnData
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Missing Data - Hist Vol, Corr")
    
    Dim baseDt As String
    baseDt = Format(ws.Range("B1").value, "yyyymmdd")
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveHistoricalVol?baseDt=" & baseDt & "&dataSetId=official"
    
    ' JSON data와 POST request를 보내는 subroutine을 호출한다.
    'SendPostRequest jsonString, url
    Set requestHandler = New CAsyncRequestHandler
    requestHandler.SendPostRequestAsync jsonString, url
End Sub

