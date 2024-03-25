Attribute VB_Name = "ClassPostVol"
Sub RunFunc()
    Dim ws As Worksheet
    Dim cell As Range
    Dim jsonString As String
    Dim dataId As String
    Dim firstObject As Boolean
    Dim postVolUpdater As postVolUpdater ' 클래스의 인스턴스를 위한 reference variable
    
    Set ws = ThisWorkbook.Sheets("Vol")
    jsonString = "["
    
    firstObject = True
    For Each cell In ws.Range("AD1:AD" & ws.Cells(ws.Rows.Count, "AD").End(xlUp).Row)
        Select Case cell.Value
            Case "KOSPI_LV"
                dataId = "KOSPI200_LOC"
            Case "NKY_LV"
                dataId = "N225_LOC"
            Case "HSI_LV"
                dataId = "HSI_LOC"
            Case "HSCEI_LV"
                dataId = "HSCEI_LOC"
            Case Else
                dataId = "" ' 값이 어떤 경우에도 맞지 않다면, skip
        End Select
        
        If dataId <> "" Then
            If Not firstObject Then
                jsonString = jsonString & ","
            End If
            
            ' 각 cell에 대해서 postVolUpdater class의 인스턴스를 생성
            Set postVolUpdater = New postVolUpdater
            With postVolUpdater
                Set .Worksheet = ws
                Set .RefCell = cell
                .dataId = dataId
            End With
            
            ' 메소드를 사용해서 jsonString을 만들어 준다.
            jsonString = jsonString & postVolUpdater.GenerateObjectJSON()
            
            firstObject = False
        End If
    Next cell
    
    jsonString = jsonString & "]"
    
    
    filePath = "C:\Users\JURO_NEW\Desktop\json_data\volData.json"
    
    fileNumber = FreeFile()
    Open filePath For Output As #fileNumber
    
    Print #fileNumber, jsonString
    
    Close #fileNumber
    ' 필요하다면 jsonString을 URLEncode하고, POST request를 한다.
    jsonString = URLEncode(jsonString)
    

    ' request에 대한 URL
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/vols?baseDt=20231228&dataSetId=TEST11"
    
    ' Send the POST request - Assuming SendPostRequest is a subroutine you have defined elsewhere
    SendPostRequest jsonString, url
End Sub

