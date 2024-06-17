Attribute VB_Name = "InputVol"
Sub Inputvol()

    Dim VoUrlBuilder As UrlBuilder
    Set VoUrlBuilder = New UrlBuilder
    
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("Market Data")
    
    Dim baseDt As String
    baseDt = Format(ws2.Range("A2").value, "yyyymmdd")
    
    VoUrlBuilder.baseurl = "http://localhost:8080/val/marketdata/"
    VoUrlBuilder.Version = "v1/"
    VoUrlBuilder.DataParameter = "vols?"
    VoUrlBuilder.baseDt = "baseDt=" & baseDt & "&"
    VoUrlBuilder.DataIds = "dataIds=HSCEI_LOC,HSI_LOC,N225_LOC,KOSPI200_LOC"
    
    Dim VoUrl As String
    VoUrl = VoUrlBuilder.MakeUrl
    
    Debug.Print VoUrl
    
    Dim JsonString As String
    JsonString = GetHttpResponseText2(VoUrl)
    
    Dim JsonResponse As Object
    Set JsonResponse = JsonConverter.ParseJson(JsonString)
    
    If JsonResponse.Exists("code") Then
        If JsonResponse("code") = "ERROR" Then
            Dim errMsg As String
            errMsg = "Error: " & JsonResponse("message")
            MsgBox errMsg, vbCritical ' Display the error message in a message box
            Exit Sub
    
        ElseIf JsonResponse("code") = "SUCCESS" Then
            Dim Volatilities As Collection
            Set Volatilities = JsonResponse("response")("volatilities")
            
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Sheets("Vol")
            
            Dim importer As New VolUpdaterNew
            With importer
                Set .Worksheet = ws
                Set .Volatilities = Volatilities
                .CodeColumn = "A"
                .ImportData
                .FillEmptyCells
                        
            End With
        End If
    End If
    
End Sub
