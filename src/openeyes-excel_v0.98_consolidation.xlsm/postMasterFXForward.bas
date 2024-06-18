Attribute VB_Name = "postMasterFXForward"
Option Explicit

Sub postMasterFXForward()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Missing Data - Fx Forward")
    
    Dim titleRange As Range
    Set titleRange = ws.Range("A:A").Find(What:="FX Forward Curve", Lookat:=xlWhole)
    
    Dim k As Long
    k = 4
    
    Dim FXForwardNames As Collection
    Set FXForwardNames = New Collection
    Dim i As Long
    For i = 1 To k
        Dim eachName As Object
        Set eachName = CreateObject("Scripting.Dictionary")
        Dim nameRange As Range
        Set nameRange = titleRange.Offset(2, 1 + 3 * (i - 1))
        
        Dim dataId As String
        dataId = nameRange.value
        
        Dim crnc As String
        Dim reltCrnc As String
        
        crnc = nameRange.Offset(2, 0).value
        reltCrnc = nameRange.Offset(1, 0).value
        
        Dim dataNM As String
        dataNM = crnc + "-" + reltCrnc + " FX Forward"
        
        eachName("dataId") = dataId
        eachName("dataNM") = dataNM
        eachName("crnc") = crnc
        eachName("reltCrnc") = reltCrnc
    
        FXForwardNames.Add eachName
    Next i
    
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(FXForwardNames)
    
    Debug.Print jsonString
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveFXForwardMaster"
    
    ' JSON data와 POST request를 보내는 subroutine을 호출한다.
    SendPostRequest jsonString, url
End Sub

