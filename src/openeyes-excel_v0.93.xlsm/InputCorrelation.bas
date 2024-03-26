Attribute VB_Name = "InputCorrelation"
Sub InputCorrelation()

    Dim corrUrlBuilder As UrlBuilder
    Set corrUrlBuilder = New UrlBuilder
    
    'setter를 이용해서 UrlBuilder의 property를 적절하게 세팅해준다.
    corrUrlBuilder.baseURL = "http://localhost:8080/val/marketdata/"
    corrUrlBuilder.Version = "v1/"
    corrUrlBuilder.DataParameter = "corrs?"
    corrUrlBuilder.baseDt = "baseDt=20231228&"
    corrUrlBuilder.DataIds = "dataIds=FXKRWHKD,HSI,HSCEI,KOSPI200,FXKRWJPY,EUROSTOXX,N225,FXKRWEUR"
    
    '메서드 이용, return값이 full url.
    Dim corrUrl As String
    corrUrl = corrUrlBuilder.MakeUrl
    
    Debug.Print corrUrl
    
    Dim jsonString As String
    jsonString = GetHttpResponseText(corrUrl)
    Debug.Print jsonString
    Dim JsonResponse As Object
    Set JsonResponse = JsonConverter.ParseJson(jsonString)
    
    Dim corrs As Collection
    Set corrs = JsonResponse("response")("correlations")
    
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("Market Data")
    
    Dim equityRow As Integer
    equityRow = Ws.Columns(1).Find(What:="Equity", LookIn:=xlValues, LookAt:=xlPart).Row
    Dim fxRow As Integer
    fxRow = Ws.Columns(1).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole).Row
    
    'Call UpdateCellsWithCorrelation(ws, SelCorrelation, equityRow + 3, equityRow + 4, "CORR", 3)
    'Call UpdateCellsWithCorrelation(ws, SelCorrelation, FXRow + 3, FXRow + 4, "CORR", 4)
    Dim corrUpdater1 As CorrUpdater
    Set corrUpdater1 = New CorrUpdater
    
    With corrUpdater1
        Set .Worksheet = Ws
        Set .SelCorrelation = corrs
        .ColumnNameRow = equityRow + 3
        .StartRow = equityRow + 4
        .MatrixId = "CORR"
        .StartColumn = 3
        .UpdateCorrelations
        
        .ColumnNameRow = fxRow + 3
        .StartRow = fxRow + 4
        .StartColumn = 4
        .UpdateCorrelations
        
    End With
    
    
    
End Sub
