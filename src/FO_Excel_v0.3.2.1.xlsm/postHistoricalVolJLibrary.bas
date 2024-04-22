Attribute VB_Name = "postHistoricalVolJLibrary"
Option Explicit

Sub postHistoricalVolModule()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Missing Data - Hist Vol, Corr")
    
    ' Create a collection to hold each row's Dictionary
    Dim dataCollection As New Collection
    
    Dim i As Long
    i = 5
    
    ' Create a Dictionary for each row and add it to the Collection
    Do While ws.Cells(i, "A").value <> ""
        Dim dataDict As Object
        Set dataDict = CreateObject("Scripting.Dictionary")
        
        ' Add key-value pairs to the Dictionary
        dataDict.Add "dataId", ws.Cells(i, "A").value & "_VOL_250"
        'dataDict("dataId") = ws.Cells(i, "A").value & "_VOL_250" (위의 것과 같은 결과를 만든다. 이 방식이 더 편해 보인다. 둘다 가능)
        dataDict.Add "histvol", ws.Cells(i, "C").value / 100
        
        ' Add the Dictionary to the Collection
        dataCollection.Add dataDict
        
        i = i + 1
    Loop
    
    ' Convert the Collection of Dictionaries to a JSON string
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(dataCollection, Whitespace:=2)
    
    ' Output the JSON string
    Debug.Print jsonString
End Sub

Sub UseHistoricalVolProcessor()
    Dim volProcessor As postHistoricalVol
    Set volProcessor = New postHistoricalVol
    
    ' Set properties
    Set volProcessor.Worksheet = ThisWorkbook.Sheets("Missing Data - Hist Vol, Corr")
    volProcessor.StartRow = 5
    
    ' Call the method to process data
    volProcessor.ProcessData
End Sub

