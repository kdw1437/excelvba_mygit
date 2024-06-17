Attribute VB_Name = "GreekInput"
Option Explicit

Sub ImportGreek()
    Dim objRequest As Object
    Dim strUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String
    Dim json As Object
    
    ' XMLHTTP object를 만든다.
    Set objRequest = CreateObject("MSXML2.XMLHTTP")

    ' JSON data를 가져오기 위해 URL을 정의한다.
    strUrl = "http://urosys-web.juroinstruments.com/app/selectGreeks?jobId=81"
    blnAsync = True

    ' GET request를 만든다.
    With objRequest
        .Open "GET", strUrl, blnAsync
        .SetRequestHeader "Content-Type", "application/json"
        .Send
        
        ' response를 기다린다.
        While .readyState <> 4
            DoEvents ' 엑셀이 GET과정 동안 freeze되지 않도록 한다.
        Wend
        
        ' response를 얻는다.
        strResponse = .ResponseText
    End With

    ' Use a JSON parser to parse the response를 parsing하기 위해 JSON parser를 이용한다.
    Set json = JsonConverter.ParseJson(strResponse)
    Dim rowNumber As Integer
    rowNumber = 4
    
    Dim colNumber As Integer
    colNumber = 2
    '이 colNumber와 rowNumber는 jsondata 바탕으로 column 생성할 때, 사용.
    '여기서는 rowNumber와 colNumber는 필요하지 않다. 왜냐하면, column명을 미리 입력해놨음.
    '칼럼명에 맞춰서 뿌려줘야 한다.
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("sheet2")
    
    'column number에 대해, JSON key의 맵핑을 만든다.
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
        ' item에서 key-value 쌍을 반복한다.
        For Each key In Item.Keys
            ' column mapping에 key가 존재하는지 확인한다.
            If colMapping.Exists(key) Then
                ' 적절한 Cell에 값을 넣는다.
                ws.Cells(rowNumber, colMapping(key)).Value = Item(key)
            End If
            DoEvents ' Excel이 어는 것을 방지한다.
        Next key
        ' 다음 row로 이동한다.
        rowNumber = rowNumber + 1
        DoEvents ' Excel이 어는 것을 방지한다.
    Next Item
    
    
    
End Sub


Sub TestImport()
    ImportGreekData "81"  ' Replace "12345" with the actual jobId you want to use
End Sub

