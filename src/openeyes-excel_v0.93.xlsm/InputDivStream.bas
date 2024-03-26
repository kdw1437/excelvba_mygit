Attribute VB_Name = "InputDivStream"
Sub InputDivStream()

    Dim discreteDivUrlBuilder As UrlBuilder
    Set discreteDivUrlBuilder = New UrlBuilder
    
    'setter를 이용해서 UrlBuilder의 property를 적절하게 세팅해준다.
    discreteDivUrlBuilder.baseURL = "http://localhost:8080/val/marketdata/"
    discreteDivUrlBuilder.Version = "v1/"
    discreteDivUrlBuilder.DataParameter = "selectDividendStream?"
    discreteDivUrlBuilder.baseDt = "baseDt=20240320&"
    discreteDivUrlBuilder.DataIds = "dataIds=KOSPI200_D,SPX_D"
    
    '메서드 이용, return값이 full url.
    Dim discreteDivUrl As String
    discreteDivUrl = discreteDivUrlBuilder.MakeUrl
    
    Dim jsonString As String
    jsonString = GetHttpResponseText(discreteDivUrl)
    
    Dim JsonResponse As Object
    Set JsonResponse = JsonConverter.ParseJson(jsonString)
    
    Dim dividendStreams As Collection
    Set dividendStreams = JsonResponse("response")("dividendStreams")
    
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("DiscreteDividend")
    
    Dim DDcell As Range
    Set DDcell = Ws.Columns(1).Find(What:="Discrete Dividend", LookIn:=xlValues, LookAt:=xlPart)
    
    If Not DDcell Is Nothing Then
        Dim StartCell As Range
        Set StartCell = DDcell.Offset(2, 0) ' 같은 column에서 DDcell로부터 2 row 밑에서 시작한다.
        
        Dim endCell As Range
        Set endCell = StartCell ' startCell로 endCell을 초기화한다.
        
        Dim cellToCheck As Range
        Set cellToCheck = StartCell ' cellToCheck를 startCell로 초기화한다.
        
        ' 한 셀 씩 건너뛰면서 체크해서 마지막 셀을 찾아낸다.
        Do While Not IsEmpty(cellToCheck.Value)
            Set endCell = cellToCheck ' endCell을 update한다.
            Set cellToCheck = cellToCheck.Offset(0, 2) ' 한 셀씩 건너뛰면서 check한다.
        Loop
        
        ' 시작 cell로 부터 끝나는 cell까지 searchRange를 정의한다.
        Dim searchRange As Range
        Set searchRange = Ws.Range(StartCell, Ws.Cells(StartCell.Row, endCell.Column))
    Else
        Debug.Print "Discrete Dividend not found."
    End If
    
    Dim divStreamUpdater As divStreamUpdater
    Set divStreamUpdater = New divStreamUpdater
    
    divStreamUpdater.SetWorksheet Ws.Name
    divStreamUpdater.SetDivStreamData dividendStreams
    divStreamUpdater.SetSearchRange searchRange
    divStreamUpdater.UpdateWorksheet
    
End Sub
