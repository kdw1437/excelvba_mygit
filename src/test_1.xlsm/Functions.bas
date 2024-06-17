Attribute VB_Name = "Functions"
Function ImportGreekData(jobId As String)
    Dim objRequest As Object
    Dim strUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String
    Dim json As Object

    ' Create the XMLHTTP object
    Set objRequest = CreateObject("MSXML2.XMLHTTP")

    ' Construct the URL with the provided jobId
    strUrl = "http://urosys-web.juroinstruments.com/app/selectGreeks?jobId=" & jobId
    blnAsync = True

    ' Open and send the GET request
    With objRequest
        .Open "GET", strUrl, blnAsync
        .SetRequestHeader "Content-Type", "application/json"
        .Send

        ' Wait for the response
        While .readyState <> 4
            DoEvents ' Prevent Excel from freezing
        Wend

        ' Get the response
        strResponse = .ResponseText
    End With

    ' Parse the JSON response
    Set json = JsonConverter.ParseJson(strResponse)
    Dim rowNumber As Integer
    rowNumber = 4

    Dim colNumber As Integer ' Not used in this function
    colNumber = 2

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("sheet2")

    ' Create column mapping
    Dim colMapping As Object
    Set colMapping = CreateObject("Scripting.Dictionary")

    colMapping("rfCd") = 8
    colMapping("jobId") = 6
    colMapping("sensTyCd") = 11
    colMapping("delta") = 9
    colMapping("itemCd") = 7
    colMapping("gamma") = 10

    Dim key As Variant
    Dim Item As Object

    ' Iterate through the JSON data
    For Each Item In json("selectGreekValues")
        For Each key In Item.Keys
            If colMapping.Exists(key) Then
                ws.Cells(rowNumber, colMapping(key)).Value = Item(key)
            End If
            DoEvents
        Next key
        rowNumber = rowNumber + 1
        DoEvents
    Next Item
End Function

