Attribute VB_Name = "InputDiscreteDiv"
Sub InputDiscreteDiv()
    Dim discreteDivUrlBuilder As UrlBuilder
    Set discreteDivUrlBuilder = New UrlBuilder
    
    'setter�� �̿��ؼ� UrlBuilder�� property�� �����ϰ� �������ش�.
    discreteDivUrlBuilder.baseURL = "http://localhost:8080/val/marketdata/"
    discreteDivUrlBuilder.Version = "v1/"
    discreteDivUrlBuilder.DataParameter = "selectDiscreteDividends?"
    discreteDivUrlBuilder.baseDt = "baseDt=20240320&"
    discreteDivUrlBuilder.DataIds = "dataIds=KOSPI200,SPX"
    
    '�޼��� �̿�, return���� full url.
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
        Set StartCell = DDcell.Offset(2, 0) ' ���� column���� DDcell�κ��� 2 row �ؿ��� �����Ѵ�.
        
        Dim endCell As Range
        Set endCell = StartCell ' startCell�� endCell�� �ʱ�ȭ�Ѵ�.
        
        Dim cellToCheck As Range
        Set cellToCheck = StartCell ' cellToCheck�� startCell�� �ʱ�ȭ�Ѵ�.
        
        ' �� �� �� �ǳʶٸ鼭 üũ�ؼ� ������ ���� ã�Ƴ���.
        Do While Not IsEmpty(cellToCheck.Value)
            Set endCell = cellToCheck ' endCell�� update�Ѵ�.
            Set cellToCheck = cellToCheck.Offset(0, 2) ' �� ���� �ǳʶٸ鼭 check�Ѵ�.
        Loop
        
        ' ���� cell�� ���� ������ cell���� searchRange�� �����Ѵ�.
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
