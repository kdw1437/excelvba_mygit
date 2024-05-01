Attribute VB_Name = "postMasterDDiv"
Option Explicit

Sub postMasterDDiv()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Missing Data - D_Dividend")
    
    Dim titleRange As Range
    Set titleRange = ws.Range("A:A").Find(What:="Discrete Dividend", Lookat:=xlWhole)

    Dim k As Long
    k = 4
    
    Dim masterDDivCollection As Collection
    Set masterDDivCollection = New Collection
    
    Dim i As Long
    
    For i = 1 To k
        
        Dim dataIdRange As Range
        Set dataIdRange = titleRange.Offset(3, 1 + 3 * (i - 1))
        
        Dim dataId As String
        dataId = dataIdRange.value
        
        Dim dataNM As String
        dataNM = dataIdRange.Offset(1, 0).value
        
        Dim crncCode As String
        crncCode = dataIdRange.Offset(2, 0).value
        
        Dim eachName As Object
        Set eachName = CreateObject("Scripting.Dictionary")
        
        eachName("dataId") = dataId
        
        eachName("crncCode") = crncCode
        
        masterDDivCollection.Add eachName
        
    Next i
    
    Dim jsonString As String
    
    jsonString = JsonConverter.ConvertToJson(masterDDivCollection)
    
    Debug.Print jsonString
    
    Dim url As String
    url = "http://localhost:8080/val/saveDisDividendMaster"
    
    ' JSON data와 POST request를 보내는 subroutine을 호출한다.
    SendPostRequest jsonString, url
End Sub
