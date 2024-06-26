Attribute VB_Name = "InputDivStream"
Sub InputDivStream()

    Dim discreteDivUrlBuilder As UrlBuilder
    Set discreteDivUrlBuilder = New UrlBuilder
    
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("DiscreteDividend")
    
    Dim baseDt As String
    baseDt = Format(ws2.Range("A2").value, "yyyymmdd")
    
    'setter를 이용해서 UrlBuilder의 property를 적절하게 세팅해준다.
    discreteDivUrlBuilder.baseURL = "http://localhost:8080/val/marketdata/"
    discreteDivUrlBuilder.Version = "v1/"
    discreteDivUrlBuilder.DataParameter = "selectDividendStream?"
    discreteDivUrlBuilder.baseDt = "baseDt=" & baseDt & "&"
    discreteDivUrlBuilder.DataIds = "dataIds=KOSPI200_D,SPX_D"
    
    '메서드 이용, return값이 full url.
    Dim discreteDivUrl As String
    discreteDivUrl = discreteDivUrlBuilder.MakeUrl
    
    Dim JsonString As String
    JsonString = GetHttpResponseText2(discreteDivUrl)
    
    Dim JsonResponse As Object
    Set JsonResponse = JsonConverter.ParseJson(JsonString)
    
    ' Check for error in the response
    If JsonResponse.Exists("code") Then
        If JsonResponse("code") = "ERROR" Then
            Dim errMsg As String
            errMsg = "Error: " & JsonResponse("message")
            MsgBox errMsg, vbCritical ' Display the error message in a message box
            Exit Sub
        
        ElseIf JsonResponse("code") = "SUCCESS" Then
            Dim dividendStreams As Collection
            Set dividendStreams = JsonResponse("response")("dividendStreams")
            
            Dim ws As Worksheet
            Set ws = ThisWorkbook.Sheets("DiscreteDividend")
            
            Dim DDcell As Range
            Set DDcell = ws.Columns(1).Find(What:="Discrete Dividend", LookIn:=xlValues, LookAt:=xlPart)
            
            If Not DDcell Is Nothing Then
                Dim startCell As Range
                Set startCell = DDcell.Offset(2, 0) ' 같은 column에서 DDcell로부터 2 row 밑에서 시작한다.
                
                Dim endCell As Range
                Set endCell = startCell ' startCell로 endCell을 초기화한다.
                
                Dim cellToCheck As Range
                Set cellToCheck = startCell ' cellToCheck를 startCell로 초기화한다.
                
                ' 한 셀 씩 건너뛰면서 체크해서 마지막 셀을 찾아낸다.
                Do While Not IsEmpty(cellToCheck.value)
                    Set endCell = cellToCheck ' endCell을 update한다.
                    Set cellToCheck = cellToCheck.Offset(0, 2) ' 한 셀씩 건너뛰면서 check한다.
                Loop
                
                ' 시작 cell로 부터 끝나는 cell까지 searchRange를 정의한다.
                Dim searchRange As Range
                Set searchRange = ws.Range(startCell, ws.Cells(startCell.row, endCell.Column))
            Else
                Debug.Print "Discrete Dividend not found."
            End If
            
            Dim divStreamUpdater As divStreamUpdater
            Set divStreamUpdater = New divStreamUpdater
            
            divStreamUpdater.SetWorksheet ws.Name
            divStreamUpdater.SetDivStreamData dividendStreams
            divStreamUpdater.SetSearchRange searchRange
            divStreamUpdater.UpdateWorksheet
        End If
        
    End If
End Sub
