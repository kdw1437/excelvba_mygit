Attribute VB_Name = "ClassPostVol"
Dim requestHandler As CAsyncRequestHandler

Sub RunFunc()
    Dim ws As Worksheet
    Dim cell As Range
    Dim jsonString As String
    Dim dataId As String
    Dim firstObject As Boolean
    Dim postVolUpdater As postVolUpdater ' Ŭ������ �ν��Ͻ��� ���� reference variable
    
    Set ws = ThisWorkbook.Sheets("Vol")
    jsonString = "["
    
    firstObject = True
    For Each cell In ws.Range("AD1:AD" & ws.Cells(ws.Rows.Count, "AD").End(xlUp).row)
        Select Case cell.value
            Case "KOSPI_LV"
                dataId = "KOSPI200_LOC"
            Case "NKY_LV"
                dataId = "N225_LOC"
            Case "HSI_LV"
                dataId = "HSI_LOC"
            Case "HSCEI_LV"
                dataId = "HSCEI_LOC"
            Case Else
                dataId = "" ' ���� � ��쿡�� ���� �ʴٸ�, skip
        End Select
        
        If dataId <> "" Then
            If Not firstObject Then
                jsonString = jsonString & ","
            End If
            
            ' �� cell�� ���ؼ� postVolUpdater class�� �ν��Ͻ��� ����
            Set postVolUpdater = New postVolUpdater
            With postVolUpdater
                Set .Worksheet = ws
                Set .RefCell = cell
                .dataId = dataId
            End With
            
            ' �޼ҵ带 ����ؼ� jsonString�� ����� �ش�.
            jsonString = jsonString & postVolUpdater.GenerateObjectJSON2()
            
            firstObject = False
        End If
        DoEvents
    Next cell
    
    jsonString = jsonString & "]"
    
    
'    filePath = "C:\Users\JURO_NEW\Desktop\json_data\volData240607.json"
'
'    fileNumber = FreeFile()
'    Open filePath For Output As #fileNumber
'
'    Print #fileNumber, JsonString
'
'    Close #fileNumber
    Debug.Print jsonString
    
    ' �ʿ��ϴٸ� jsonString�� URLEncode�ϰ�, POST request�� �Ѵ�.
    jsonString = URLEncode(jsonString)
    
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("Market Data")
    
    Dim dataSetId As String
    dataSetId = ws2.Range("O2").value
    
    Dim baseDt As String
    baseDt = Format(ws2.Range("A2").value, "yyyymmdd")
    ' request�� ���� URL
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveVols?baseDt=" & baseDt & "&dataSetId=" & dataSetId
    
    ' Send the POST request - Assuming SendPostRequest is a subroutine you have defined elsewhere
'     SendPostRequest JsonString, url
    Set requestHandler = New CAsyncRequestHandler
    requestHandler.SendPostRequestAsync jsonString, url
    
End Sub

