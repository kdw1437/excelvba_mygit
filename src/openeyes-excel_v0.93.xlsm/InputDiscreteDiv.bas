Attribute VB_Name = "InputDiscreteDiv"
Sub InputDiscreteDiv()
    Dim discreteDivUrlBuilder As UrlBuilder
    Set discreteDivUrlBuilder = New UrlBuilder
    
    'setter를 이용해서 UrlBuilder의 property를 적절하게 세팅해준다.
    discreteDivUrlBuilder.baseURL = "http://localhost:8080/val/marketdata/"
    discreteDivUrlBuilder.Version = "v1/"
    discreteDivUrlBuilder.DataParameter = "selectDiscreteDividends?"
    discreteDivUrlBuilder.baseDt = "baseDt=20240320&"
    discreteDivUrlBuilder.DataIds = "dataIds=KOSPI200,SPX"
    
    '메서드 이용, return값이 full url.
    Dim discreteDivUrl As String
    discreteDivUrl = discreteDivUrlBuilder.MakeUrl
    
    Dim jsonString As String
    jsonString = GetHttpResponseText(discreteDivUrl)
    
    Dim JsonResponse As Object
    Set JsonResponse = JsonConverter.ParseJson(jsonString)
    
    Dim discreteDiv As Collection
    Set discreteDiv = JsonResponse("response")("discreteDividendCurves")
    
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
    
    Dim i As Integer
    Dim j As Integer
    Dim dataSet As Object
    Dim dividend As Object
    Dim targetRow As Range
    Dim dataId As String
    Dim searchCell As Range
    Dim findCell As Range
    Dim dataValue As Object
    Dim divValue As Object
    
    For i = 1 To discreteDiv.Count
        Set dataSet = discreteDiv(i)
        dataId = dataSet("dataId")
        
        For Each searchCell In searchRange
            If searchCell.Value = dataId Then
                Set findCell = searchCell
            End If
        Next searchCell
            
        If findCell.Value = dataSet("dataId") Then
            For j = 1 To dataSet("discreteDividends").Count
                Set divValue = dataSet("discreteDividends")(j)
                findCell.Offset(j + 1, 0).Value = divValue("date")
                findCell.Offset(j + 1, 1).Value = divValue("value")
                
            Next j
        End If
        
               
    Next i
    
    
End Sub
