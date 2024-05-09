Attribute VB_Name = "ClassPostCorrhardcoded"
Sub PrintJsonString()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")

    ' PostCorrUpdater class�� �ν��Ͻ� ����
    Dim postCorrUpdater As New postCorrUpdater
    With postCorrUpdater
        ' vertical ranget�� horizontal range setter ȣ��
        Set .VerticalRange = ws.Range("M8", ws.Range("M8").End(xlDown))
        Set .HorizontalRange = ws.Range("O7", ws.Range("O7").End(xlToRight))
        
        
        'Debug.Print .CorrJson()
    End With
    Dim DataString As String
    DataString = postCorrUpdater.CorrJsonv2()
    Debug.Print DataString
    DataString = URLEncode(DataString)
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/corrs?baseDt=20231228&dataSetId=TEST11&matrixId=CORR"
    
    ' JSON data�� POST request�� ������ ���� subroutine�� ȣ���Ѵ�.
    SendPostRequest DataString, url
End Sub

Sub PrintJsonString2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")
    
    Dim postCorrUpdater As New postCorrUpdater
    With postCorrUpdater
    
        Set .VerticalRange = ws.Range("M20", ws.Range("M20").End(xlDown))
        Set .HorizontalRange = ws.Range("P19", ws.Range("P19").End(xlToRight))
        
    End With
    Dim DataString As String
    DataString = postCorrUpdater.CorrJsonh2()
    Debug.Print DataString
    DataString = URLEncode(DataString)
    
    
        
        
        
End Sub
