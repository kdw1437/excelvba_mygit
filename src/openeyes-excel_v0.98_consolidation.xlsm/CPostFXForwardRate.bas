Attribute VB_Name = "CPostFXForwardRate"
Option Explicit
Dim requestHandler As CAsyncRequestHandler

Sub UseFXForwardData()
    Dim fxData As New PostForwardRate
    
    Set fxData.Worksheet = ThisWorkbook.Sheets("Missing Data - Fx Forward")
    Set fxData.StartCell = fxData.Worksheet.Range("A:A").Find(What:="FX Forward Curve", Lookat:=xlWhole)
    fxData.k = 4  ' Setting how many currencies to process
    
    'fxData.GenerateJSON
    
    Dim jsonString As String
    jsonString = fxData.ReturnJSON2
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Missing Data - Fx Forward")
    
    Dim baseDt As String
    baseDt = Format(ws.Range("B1").value, "yyyymmdd")
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveForwardFX?baseDt=" & baseDt & "&dataSetId=official"
    
    ' JSON data와 POST request를 보내는 subroutine을 호출한다.
    'SendPostRequest jsonString, url
    Set requestHandler = New CAsyncRequestHandler
    requestHandler.SendPostRequestAsync jsonString, url
End Sub


