Attribute VB_Name = "CPostDisDiv"
Option Explicit
Dim requestHandler As CAsyncRequestHandler

Sub UseDividendDataProcessor()
    Dim dp As postDisDivMissingData
    Set dp = New postDisDivMissingData
    
    Set dp.Worksheet = ThisWorkbook.Worksheets("Missing Data - D_Dividend(����)")
    Set dp.StartCell = dp.Worksheet.Range("A:A").Find(What:="Discrete Dividend", Lookat:=xlWhole)
    dp.k = 4 ' (K�� �˷��� ���̰� dynamic�ϰ� �����Ǿ��� ���� �ű⿡ ���缭 �ڵ� �ۼ�)
    
    If Not dp.StartCell Is Nothing Then
        Dim jsonString As String
        
        jsonString = dp.ReturnJSON
        
        Debug.Print jsonString
    Else
        MsgBox "Start cell not found."
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Missing Data - D_Dividend(����)")
    
    Dim baseDt As String
    baseDt = Format(ws.Range("B1").value, "yyyymmdd")
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveDividendStream?baseDt=" & baseDt & "&dataSetId=official"
    
    ' JSON data�� POST request�� ������ subroutine�� ȣ���Ѵ�.
    'SendPostRequest jsonString, url
    Set requestHandler = New CAsyncRequestHandler
    requestHandler.SendPostRequestAsync jsonString, url
End Sub


