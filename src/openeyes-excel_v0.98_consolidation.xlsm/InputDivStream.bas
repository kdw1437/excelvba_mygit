Attribute VB_Name = "InputDivStream"
Sub InputDivStream()

    Dim discreteDivUrlBuilder As UrlBuilder
    Set discreteDivUrlBuilder = New UrlBuilder
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DiscreteDividend")
    
    Dim baseDt As String
    baseDt = Format(ws.Range("A2").value, "yyyymmdd")
    
    'setter�� �̿��ؼ� UrlBuilder�� property�� �����ϰ� �������ش�.
    discreteDivUrlBuilder.baseurl = "http://localhost:8080/val/marketdata/"
    discreteDivUrlBuilder.Version = "v1/"
    discreteDivUrlBuilder.DataParameter = "selectDividendStream?"
    discreteDivUrlBuilder.baseDt = "baseDt=" & baseDt & "&"
    'discreteDivUrlBuilder.dataIds = "dataIds=KOSPI200_D,SPX_D"
    
    ' column A���� "Discrete Dividend" cell�� ã�´�.
    Dim DDcell As Range
    Set DDcell = ws.Columns(1).Find(What:="Discrete Dividend", LookIn:=xlValues, LookAt:=xlPart)
    
    Dim dataIds As String
    If Not DDcell Is Nothing Then
        ' "Discrete Dividend" cell���� 2 row �Ʒ����� �����Ѵ�.
        Dim startCell As Range
        Set startCell = DDcell.Offset(2, 0)
        
        ' empty cell�� ã���� �� ����, 2 column�� �̵��Ѵ�.
        Dim currentCell As Range
        Set currentCell = startCell
        
        Do While Not IsEmpty(currentCell.value)
            If dataIds <> "" Then
                dataIds = dataIds & ","
            End If
            dataIds = dataIds & currentCell.value
            Set currentCell = currentCell.Offset(0, 2)
        Loop
    End If
    
    discreteDivUrlBuilder.dataIds = "dataIds=" & dataIds
    '�޼��� �̿�, return���� full url.
    Dim discreteDivUrl As String
    discreteDivUrl = discreteDivUrlBuilder.MakeUrl
    
    Dim jsonString As String
    jsonString = GetHttpResponseText2(discreteDivUrl)
    
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
            Dim dividendStreams As Collection
            Set dividendStreams = JsonResponse("response")("dividendStreams")
            
            If Not DDcell Is Nothing Then
                
                Set startCell = DDcell.Offset(2, 0) ' ���� column���� DDcell�κ��� 2 row �ؿ��� �����Ѵ�.
                
                Dim endCell As Range
                Set endCell = startCell ' startCell�� endCell�� �ʱ�ȭ�Ѵ�.
                
                Dim cellToCheck As Range
                Set cellToCheck = startCell ' cellToCheck�� startCell�� �ʱ�ȭ�Ѵ�.
                
                ' �� �� �� �ǳʶٸ鼭 üũ�ؼ� ������ ���� ã�Ƴ���.
                Do While Not IsEmpty(cellToCheck.value)
                    Set endCell = cellToCheck ' endCell�� update�Ѵ�.
                    Set cellToCheck = cellToCheck.Offset(0, 2) ' �� ���� �ǳʶٸ鼭 check�Ѵ�.
                Loop
                
                ' ���� cell�� ���� ������ cell���� searchRange�� �����Ѵ�.
                Dim searchRange As Range
                Set searchRange = ws.Range(startCell, ws.Cells(startCell.row, endCell.Column))
            Else
                Debug.Print "Discrete Dividend not found."
            End If
            
            Dim divStreamUpdater As divStreamUpdater
            Set divStreamUpdater = New divStreamUpdater
            
            divStreamUpdater.SetWorksheet ws.name
            divStreamUpdater.SetDivStreamData dividendStreams
            divStreamUpdater.SetSearchRange searchRange
            divStreamUpdater.UpdateWorksheet
        End If
        
    End If
End Sub
