Attribute VB_Name = "postCorrmodule"
Option Explicit

Sub postCorrModule()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Missing Data - Hist Vol, Corr")
    
    Dim jsonString As String
    jsonString = "["
    
    Dim i As Long
    i = 5
    
    ' Loop until an empty cell is found in column E
    Do While ws.Cells(i, "E").value <> ""
        ' Construct the JSON string
        jsonString = jsonString & "{""dataId"":""" & ws.Cells(i, "E").value & ":" & ws.Cells(i, "F").value & ""","
        jsonString = jsonString & """dataId1"":""" & ws.Cells(i, "E").value & ""","
        jsonString = jsonString & """dataId2"":""" & ws.Cells(i, "F").value & ""","
        jsonString = jsonString & """corr"":" & ws.Cells(i, "G").value & "}"
        
        i = i + 1 ' Move to the next row
        
        ' If the next pair of dataId1 or dataId2 is not empty, add a comma to separate the objects
        If ws.Cells(i, "E").value <> "" Or ws.Cells(i, "F").value <> "" Then
            jsonString = jsonString & ","
        End If
    Loop
    
    jsonString = jsonString & "]"
    
    ' Output the jsonString to Immediate Window (Ctrl + G in VBA Editor to view)
    Debug.Print jsonString
    

End Sub
