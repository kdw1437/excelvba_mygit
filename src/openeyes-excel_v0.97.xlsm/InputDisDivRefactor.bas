Attribute VB_Name = "InputDisDivRefactor"
Sub InputDisDivRefactor()
    Dim discreteDivUrlBuilder As UrlBuilder
    Set discreteDivUrlBuilder = New UrlBuilder
    
    'setter�� �̿��ؼ� UrlBuilder�� property�� �����ϰ� �������ش�.
    discreteDivUrlBuilder.baseURL = "http://localhost:8080/val/marketdata/"
    discreteDivUrlBuilder.Version = "v1/"
    discreteDivUrlBuilder.DataParameter = "selectDiscreteDividends?"
    discreteDivUrlBuilder.baseDt = "baseDt=20240320&"
    discreteDivUrlBuilder.DataIds = "dataIds=KOSPI200_D,SPX_D"
    
    '�޼��� �̿�, return���� full url.
    Dim discreteDivUrl As String
    discreteDivUrl = discreteDivUrlBuilder.MakeUrl
    
    Dim JsonString As String
    JsonString = GetHttpResponseText(discreteDivUrl)
    
    Dim JsonResponse As Object
    Set JsonResponse = JsonConverter.ParseJson(JsonString)
    
    Dim discreteDiv As Collection
    Set discreteDiv = JsonResponse("response")("discreteDividendCurves")
    
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("DiscreteDividend")
    
    Dim DDcell As Range
    Set DDcell = Ws.Columns(1).Find(What:="Discrete Dividend", LookIn:=xlValues, LookAt:=xlPart)
    
'    If DDcell Is Nothing Then
'        Debug.Print "Discrete Dividend not found."
'        Exit Sub
'    End If
    
    ' searchRange ����
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
        Set searchRange = Ws.Range(startCell, Ws.Cells(startCell.row, endCell.Column))
    Else
        Debug.Print "Discrete Dividend not found."
    End If
    
    
    Dim i As Integer, j As Integer
    Dim dataSet As Object, divValue As Object
    Dim findCell As Range
    
    For i = 1 To discreteDiv.Count
        Set dataSet = discreteDiv(i)
        Set findCell = Nothing ' ���ο� dataSet�� ���ؼ� findCell�� ã�´�.
        
        ' Find the cell with the matching dataId
        For Each findCell In searchRange.Cells
            If findCell.value = dataSet("dataId") Then Exit For
            Set findCell = Nothing ' ã������ ������, findCell�� Nothing�̴�.
        Next findCell
        
        ' cell�� ã����, data�� cell�� ä���.
        If Not findCell Is Nothing Then
            For j = 1 To dataSet("discreteDividends").Count
                Set divValue = dataSet("discreteDividends")(j)
                findCell.Offset(j + 1, 0).value = divValue("date")
                findCell.Offset(j + 1, 1).value = divValue("value")
            Next j
        Else
            Debug.Print "DataId " & dataSet("dataId") & " not found in the searchRange."
        End If
    Next i
    
End Sub
