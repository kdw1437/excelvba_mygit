Attribute VB_Name = "InputVol"
Sub Inputvol()

    Dim VoUrlBuilder As UrlBuilder
    Set VoUrlBuilder = New UrlBuilder
    
    VoUrlBuilder.baseURL = "http://localhost:8080/val/marketdata/"
    VoUrlBuilder.Version = "v1/"
    VoUrlBuilder.DataParameter = "vols?"
    VoUrlBuilder.baseDt = "baseDt=20231228&"
    VoUrlBuilder.DataIds = "dataIds=HSCEI_LOC,HSI_LOC,N225_LOC,KOSPI200_LOC"
    
    Dim VoUrl As String
    VoUrl = VoUrlBuilder.MakeUrl
    
    Debug.Print VoUrl
    
    Dim JsonResponse As Object
    Set JsonResponse = GetJsonResponse(VoUrl)
    Dim Volatilities As Collection
    Set Volatilities = JsonResponse("response")("volatilities")
    
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("Vol")
    
    Dim importer As New VolUpdaterNew
    With importer
        Set .Worksheet = Ws
        Set .Volatilities = Volatilities
        .CodeColumn = "A"
        .ImportData
        .FillEmptyCells
                
    End With
    
    
End Sub
