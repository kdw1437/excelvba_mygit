Attribute VB_Name = "InputCorrelation"
Sub InputCorrelation()

    Dim corrUrlBuilder As UrlBuilder
    Set corrUrlBuilder = New UrlBuilder
    
    'setter�� �̿��ؼ� UrlBuilder�� property�� �����ϰ� �������ش�.
    corrUrlBuilder.baseURL = "http://localhost:8080/val/marketdata/"
    corrUrlBuilder.Version = "v1/"
    corrUrlBuilder.DataParameter = "corrs?"
    corrUrlBuilder.baseDt = "baseDt=20231228&"
    corrUrlBuilder.DataIds = "dataIds=FXKRWHKD,HSI,HSCEI,KOSPI200,FXKRWJPY,EUROSTOXX,N225,FXKRWEUR"
    
    '�޼��� �̿�, return���� full url.
    Dim corrUrl As String
    corrUrl = corrUrlBuilder.MakeUrl
    
    Debug.Print corrUrl
    
    Dim JsonString As String
    JsonString = GetHttpResponseText(corrUrl)
    Debug.Print JsonString
    Dim JsonResponse As Object
    Set JsonResponse = JsonConverter.ParseJson(JsonString)
    
    Dim corrs As Collection
    Set corrs = JsonResponse("response")("correlations")
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    
    Dim equityRow As Integer
    equityRow = ws.Columns(1).Find(What:="Equity", LookIn:=xlValues, LookAt:=xlPart).row
    Dim fxRow As Integer
    fxRow = ws.Columns(1).Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole).row
    
    'Call UpdateCellsWithCorrelation(ws, SelCorrelation, equityRow + 3, equityRow + 4, "CORR", 3)
    'Call UpdateCellsWithCorrelation(ws, SelCorrelation, FXRow + 3, FXRow + 4, "CORR", 4)
    Dim corrUpdater1 As CorrUpdater
    Set corrUpdater1 = New CorrUpdater
    
    With corrUpdater1
        Set .Worksheet = ws
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
