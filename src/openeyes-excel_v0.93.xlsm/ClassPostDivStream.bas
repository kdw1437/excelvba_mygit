Attribute VB_Name = "ClassPostDivStream"
Option Explicit

Sub ClassPostDivStream()

    Dim divStreamUpdater As PostDivStreamUpdater
    Set divStreamUpdater = New PostDivStreamUpdater
    
    Dim Wst As Worksheet
    Set Wst = ThisWorkbook.Sheets("DiscreteDividend")
    
    Set divStreamUpdater.Ws = Wst
    Set divStreamUpdater.TitleCell = divStreamUpdater.Ws.Cells(3, 10)
    Set divStreamUpdater.StartCell = divStreamUpdater.TitleCell.Offset(2, 0)
    
    divStreamUpdater.PopulateArrayOfIndex
    
    Dim DataString As String
    DataString = divStreamUpdater.GenerateDataString
        
    Debug.Print DataString
    
    
End Sub
