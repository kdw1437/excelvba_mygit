Attribute VB_Name = "InputCorrelation"
Sub InputCorrelation()

    Dim corrUrlBuilder As UrlBuilder
    Set corrUrlBuilder = New UrlBuilder
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    
    Dim baseDt As String
    baseDt = Format(ws.Range("A2").value, "yyyymmdd")
    'setter를 이용해서 UrlBuilder의 property를 적절하게 세팅해준다.
    corrUrlBuilder.baseurl = "http://localhost:8080/val/marketdata/"
    corrUrlBuilder.Version = "v1/"
    corrUrlBuilder.DataParameter = "corrs?"
    
    corrUrlBuilder.baseDt = "baseDt=" & baseDt & "&"
    'corrUrlBuilder.baseDt = "baseDt=20231228&"
    'corrUrlBuilder.dataIds = "dataIds=FXKRWHKD,HSI,HSCEI,KOSPI200,FXKRWJPY,EUROSTOXX,N225,FXKRWEUR"
    
    ' dataIds를 Equity cell로 부터 생성한다.
    Dim equityCell As Range
    Set equityCell = ws.Columns("A").Find(What:="Equity", LookIn:=xlValues, LookAt:=xlWhole)
    
    Dim dataIds As String
    If Not equityCell Is Nothing Then
        Dim equityDataIdsRange As Range
        Dim equityDataIdsCell As Range
        
        Set equityDataIdsRange = ws.Range(equityCell.Offset(4, 0), equityCell.Offset(4, 0).End(xlDown))
        
        For Each equityDataIdsCell In equityDataIdsRange
            If dataIds <> "" Then
                dataIds = dataIds & ","
            End If
            dataIds = dataIds & equityDataIdsCell.value
        Next equityDataIdsCell
    End If
    
    ' FX cell로부터 dataIds를 생성한다.
    Dim fxCell As Range
    Set fxCell = ws.Columns("A").Find(What:="FX", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not fxCell Is Nothing Then
        Dim fxDataIdsRange As Range
        Dim fxDataIdsCell As Range
        
        Set fxDataIdsRange = ws.Range(fxCell.Offset(4, 0), fxCell.Offset(4, 0).End(xlDown))
        
        For Each fxDataIdsCell In fxDataIdsRange
            If dataIds <> "" Then
                dataIds = dataIds & ","
            End If
            dataIds = dataIds & fxDataIdsCell.value
        Next fxDataIdsCell
    End If
    
    corrUrlBuilder.dataIds = "dataIds=" & dataIds
    
    '메서드 이용, return값이 full url.
    Dim corrUrl As String
    corrUrl = corrUrlBuilder.MakeUrl
    
    Debug.Print corrUrl
    
    Dim jsonString As String
    jsonString = GetHttpResponseText2(corrUrl)
    Debug.Print jsonString
    Dim JsonResponse As Object
    Set JsonResponse = JsonConverter.ParseJson(jsonString)
    
    ' Check for error in the response
    If JsonResponse.Exists("code") Then
        If JsonResponse("code") = "ERROR" Then
            Dim errMsg As String
            errMsg = "Error: " & JsonResponse("message")
            MsgBox errMsg, vbCritical ' Display the error message in a message box
            Exit Sub
        
        ElseIf JsonResponse("code") = "SUCCESS" Then
        'SUCCESS이면, correlation data를 처리한다.
        ' Process the correlation data if no error
            Dim corrs As Collection
            Set corrs = JsonResponse("response")("correlations")
            ' Additional code to process correlations can be added here
        
        
    '    Dim corrs As Collection
    '    Set corrs = JsonResponse("response")("correlations")
        
'            Dim ws As Worksheet
'            Set ws = ThisWorkbook.Sheets("Market Data")
            
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
                .startRow = equityRow + 4
                .MatrixId = "CORR"
                .StartColumn = 3
                .UpdateCorrelations
                
                .ColumnNameRow = fxRow + 3
                .startRow = fxRow + 4
                .StartColumn = 4
                .UpdateCorrelations
                
            End With
        End If
    End If
    
End Sub
