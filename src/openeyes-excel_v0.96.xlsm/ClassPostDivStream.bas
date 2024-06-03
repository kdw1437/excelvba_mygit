Attribute VB_Name = "ClassPostDivStream"
Option Explicit

Sub ClassPostDivStream()

    Dim divStreamUpdater As PostDivStreamUpdater
    Set divStreamUpdater = New PostDivStreamUpdater
    
    Dim Wst As Worksheet
    Set Wst = ThisWorkbook.Sheets("DiscreteDividend")
    
    Set divStreamUpdater.ws = Wst
    Set divStreamUpdater.TitleCell = divStreamUpdater.ws.Cells(3, 10)
    Set divStreamUpdater.startCell = divStreamUpdater.TitleCell.Offset(2, 0)
    
    divStreamUpdater.PopulateArrayOfIndex
    
    Dim DataString As String
    DataString = divStreamUpdater.GenerateDataString2()
        
    Debug.Print DataString
    
    DataString = URLEncode(DataString)

    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveDividendStream?baseDt=20231228&dataSetId=TEST2"

    ' POST request를 보낸다.
    SendPostRequest DataString, url
End Sub
