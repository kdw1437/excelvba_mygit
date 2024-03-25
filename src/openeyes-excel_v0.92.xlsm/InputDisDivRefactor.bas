Attribute VB_Name = "InputDisDivRefactor"
Sub InputDisDivRefactor()
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
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DiscreteDividend")
    
    Dim DDcell As Range
    Set DDcell = ws.Columns(1).Find(What:="Discrete Dividend", LookIn:=xlValues, LookAt:=xlPart)
    
    If DDcell Is Nothing Then
        Debug.Print "Discrete Dividend not found."
        Exit Sub
    End If
    
    Dim startCell As Range
    Set startCell = DDcell.Offset(2, 0) ' DDcell�κ��� 2ĭ �Ʒ����� ����
    
    ' searchRange ����
    Dim searchRange As Range
    Set searchRange = ws.Range(startCell, ws.Cells(startCell.Row, ws.Columns.Count).End(xlToLeft))
    
    Dim i As Integer, j As Integer
    Dim dataSet As Object, divValue As Object
    Dim findCell As Range
    
    For i = 1 To discreteDiv.Count
        Set dataSet = discreteDiv(i)
        Set findCell = Nothing ' ���ο� dataSet�� ���ؼ� findCell�� ã�´�.
        
        ' Find the cell with the matching dataId
        For Each findCell In searchRange.Cells
            If findCell.Value = dataSet("dataId") Then Exit For
            Set findCell = Nothing ' ã������ ������, findCell�� Nothing�̴�.
        Next findCell
        
        ' cell�� ã����, data�� cell�� ä���.
        If Not findCell Is Nothing Then
            For j = 1 To dataSet("discreteDividends").Count
                Set divValue = dataSet("discreteDividends")(j)
                findCell.Offset(j + 1, 0).Value = divValue("date")
                findCell.Offset(j + 1, 1).Value = divValue("value")
            Next j
        Else
            Debug.Print "DataId " & dataSet("dataId") & " not found in the searchRange."
        End If
    Next i
    
End Sub

