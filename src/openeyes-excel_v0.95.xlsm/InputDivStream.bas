Attribute VB_Name = "InputDivStream"
Sub InputDivStream()

    Dim discreteDivUrlBuilder As UrlBuilder
    Set discreteDivUrlBuilder = New UrlBuilder
    
    'setter�� �̿��ؼ� UrlBuilder�� property�� �����ϰ� �������ش�.
    discreteDivUrlBuilder.baseURL = "http://localhost:8080/val/marketdata/"
    discreteDivUrlBuilder.Version = "v1/"
    discreteDivUrlBuilder.DataParameter = "selectDividendStream?"
    discreteDivUrlBuilder.baseDt = "baseDt=20240320&"
    discreteDivUrlBuilder.DataIds = "dataIds=KOSPI200_D,SPX_D"
    
    '�޼��� �̿�, return���� full url.
    Dim discreteDivUrl As String
    discreteDivUrl = discreteDivUrlBuilder.MakeUrl
    
    Dim jsonString As String
    jsonString = GetHttpResponseText(discreteDivUrl)
    
    Dim JsonResponse As Object
    Set JsonResponse = JsonConverter.ParseJson(jsonString)
    
    Dim dividendStreams As Collection
    Set dividendStreams = JsonResponse("response")("dividendStreams")
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DiscreteDividend")
    
    Dim DDcell As Range
    Set DDcell = ws.Columns(1).Find(What:="Discrete Dividend", LookIn:=xlValues, LookAt:=xlPart)
    
    If Not DDcell Is Nothing Then
        Dim startCell As Range
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
    
    divStreamUpdater.SetWorksheet ws.Name
    divStreamUpdater.SetDivStreamData dividendStreams
    divStreamUpdater.SetSearchRange searchRange
    divStreamUpdater.UpdateWorksheet
    
End Sub
