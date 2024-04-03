Attribute VB_Name = "ClassPostCorrhardcoded"
Sub PrintJsonString()
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("Market Data")

    ' PostCorrUpdater class�� �ν��Ͻ� ����
    Dim postCorrUpdater As New postCorrUpdater
    With postCorrUpdater
        ' vertical ranget�� horizontal range setter ȣ��
        Set .VerticalRange = Ws.Range("M8", Ws.Range("M8").End(xlDown))
        Set .HorizontalRange = Ws.Range("O7", Ws.Range("O7").End(xlToRight))
        
        
        'Debug.Print .CorrJson()
    End With
    Dim DataString As String
    DataString = postCorrUpdater.CorrJsonv()
    Debug.Print DataString
    DataString = URLEncode(DataString)
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/corrs?baseDt=20231228&dataSetId=TEST11&matrixId=CORR"
    
    ' JSON data�� POST request�� ������ ���� subroutine�� ȣ���Ѵ�.
    SendPostRequest DataString, url
End Sub

Sub PrintJsonString2()
    Dim Ws As Worksheet
    Set Ws = ThisWorkbook.Sheets("Market Data")
    
    Dim postCorrUpdater As New postCorrUpdater
    With postCorrUpdater
    
        Set .VerticalRange = Ws.Range("M20", Ws.Range("M20").End(xlDown))
        Set .HorizontalRange = Ws.Range("P19", Ws.Range("P19").End(xlToRight))
        
    End With
    Dim DataString As String
    DataString = postCorrUpdater.CorrJsonh()
    Debug.Print DataString
    DataString = URLEncode(DataString)
    
    
        
        
        
End Sub
