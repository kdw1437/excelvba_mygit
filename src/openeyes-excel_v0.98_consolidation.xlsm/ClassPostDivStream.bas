Attribute VB_Name = "ClassPostDivStream"
Option Explicit

Dim requestHandler As CAsyncRequestHandler


Sub ClassPostDivStream()

    Dim divStreamUpdater As PostDivStreamUpdater
    Set divStreamUpdater = New PostDivStreamUpdater
    
    Dim Wst As Worksheet
    Set Wst = ThisWorkbook.Sheets("DiscreteDividend")
    
    Set divStreamUpdater.ws = Wst
    Set divStreamUpdater.TitleCell = divStreamUpdater.ws.Cells(3, 10)
    Set divStreamUpdater.StartCell = divStreamUpdater.TitleCell.Offset(2, 0)
    
    divStreamUpdater.PopulateArrayOfIndex
    
    Dim DataString As String
    DataString = divStreamUpdater.GenerateDataString2()
        
    Debug.Print DataString
    
    DataString = URLEncode(DataString)

    Dim dataSetId As String
    dataSetId = Wst.Range("J2").value
    
    Dim baseDt As String
    baseDt = Format(Wst.Range("A2").value, "yyyymmdd")
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveDividendStream?baseDt=" & baseDt & "&dataSetId=" & dataSetId
    

    ' POST request를 보낸다.
    'SendPostRequest DataString, url
    Set requestHandler = New CAsyncRequestHandler
    ' POST request를 보낸다.
    
    requestHandler.SendPostRequestAsync DataString, url
End Sub
