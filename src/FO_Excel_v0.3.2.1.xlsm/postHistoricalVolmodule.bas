Attribute VB_Name = "postHistoricalVolmodule"
Option Explicit

Sub postHistoricalVolModule()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Missing Data - Hist Vol, Corr")
    
    Dim jsonString As String
    jsonString = "["
    
    Dim i As Long
    i = 5
    
    Do While ws.Cells(i, "A").value <> ""
        jsonString = jsonString & "{""dataId"":""" & ws.Cells(i, "A").value & "_VOL_250"", ""histvol"":" & ws.Cells(i, "C").value / 100 & "}"
        i = i + 1
        If ws.Cells(i, "A").value <> "" Then
            jsonString = jsonString & ","
        End If
    Loop
    
    jsonString = jsonString & "]"
    
    Debug.Print jsonString

End Sub
