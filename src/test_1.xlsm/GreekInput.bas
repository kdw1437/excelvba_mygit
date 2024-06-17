Attribute VB_Name = "GreekInput"
Option Explicit

Sub ImportGreek()
    Dim objRequest As Object
    Dim strUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String
    Dim json As Object
    
    ' XMLHTTP object�� �����.
    Set objRequest = CreateObject("MSXML2.XMLHTTP")

    ' JSON data�� �������� ���� URL�� �����Ѵ�.
    strUrl = "http://urosys-web.juroinstruments.com/app/selectGreeks?jobId=81"
    blnAsync = True

    ' GET request�� �����.
    With objRequest
        .Open "GET", strUrl, blnAsync
        .SetRequestHeader "Content-Type", "application/json"
        .Send
        
        ' response�� ��ٸ���.
        While .readyState <> 4
            DoEvents ' ������ GET���� ���� freeze���� �ʵ��� �Ѵ�.
        Wend
        
        ' response�� ��´�.
        strResponse = .ResponseText
    End With

    ' Use a JSON parser to parse the response�� parsing�ϱ� ���� JSON parser�� �̿��Ѵ�.
    Set json = JsonConverter.ParseJson(strResponse)
    Dim rowNumber As Integer
    rowNumber = 4
    
    Dim colNumber As Integer
    colNumber = 2
    '�� colNumber�� rowNumber�� jsondata �������� column ������ ��, ���.
    '���⼭�� rowNumber�� colNumber�� �ʿ����� �ʴ�. �ֳ��ϸ�, column���� �̸� �Է��س���.
    'Į���� ���缭 �ѷ���� �Ѵ�.
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("sheet2")
    
    'column number�� ����, JSON key�� ������ �����.
    Dim colMapping As Object
    Set colMapping = CreateObject("Scripting.Dictionary")
    
    colMapping("rfCd") = 4
    colMapping("jobId") = 2
    colMapping("sensTyCd") = 7
    colMapping("delta") = 5
    colMapping("itemCd") = 3
    colMapping("gamma") = 6

    Dim key As Variant
    Dim Item As Object
    
    For Each Item In json("selectGreekValues")
        ' item���� key-value ���� �ݺ��Ѵ�.
        For Each key In Item.Keys
            ' column mapping�� key�� �����ϴ��� Ȯ���Ѵ�.
            If colMapping.Exists(key) Then
                ' ������ Cell�� ���� �ִ´�.
                ws.Cells(rowNumber, colMapping(key)).Value = Item(key)
            End If
            DoEvents ' Excel�� ��� ���� �����Ѵ�.
        Next key
        ' ���� row�� �̵��Ѵ�.
        rowNumber = rowNumber + 1
        DoEvents ' Excel�� ��� ���� �����Ѵ�.
    Next Item
    
    
    
End Sub


Sub TestImport()
    ImportGreekData "81"  ' Replace "12345" with the actual jobId you want to use
End Sub

