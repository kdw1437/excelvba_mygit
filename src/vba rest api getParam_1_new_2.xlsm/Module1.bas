Attribute VB_Name = "Module1"
Sub ImportJSONtoExcel()
    Dim objRequest As Object
    Dim strUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String
    Dim json As Object
    Dim item As Object
    Dim rowNumber As Long
    Dim colNumber As Integer

    ' Create an XMLHTTP object.
    Set objRequest = CreateObject("MSXML2.XMLHTTP")

    ' Define the URL to fetch JSON data.
    strUrl = "http://localhost:8080/val/getParam"
    blnAsync = True

    ' Create a GET request.
    With objRequest
        .Open "GET", strUrl, blnAsync
        .SetRequestHeader "Content-Type", "application/json"
        .Send
        
        ' Wait for the response.
        While .readyState <> 4
            DoEvents ' Allow Excel to process other events.
        Wend
        
        ' Get the response.
        strResponse = .ResponseText
    End With

    ' Use a JSON parser to parse the response.
    Set json = JsonConverter.ParseJson(strResponse)

    ' Start writing data from row 9 in the worksheet (B9).
    rowNumber = 9

    ' Start data input from column B (column index 2).
    colNumber = 10

    ' Write headers (keys from JSON data) to Excel.
    For Each Key In json("getParam1")(1).Keys
        Cells(rowNumber, colNumber).Value = Key
        colNumber = colNumber + 1
        DoEvents ' Prevent Excel from freezing.
    Next Key

    ' Reset the column counter to 2 (for column B) and increment the row number to start data writing.
    colNumber = 10
    rowNumber = rowNumber + 1

    ' Iterate over each item in the getParam2 array.
    For Each item In json("getParam1")
        ' Reset the column counter for each row (for column B).
        colNumber = 10
        ' Iterate over each key-value pair in the item.
        For Each Key In item.Keys
            ' Write the value to the Excel cell.
            Cells(rowNumber, colNumber).Value = item(Key)
            ' Move to the next column.
            colNumber = colNumber + 1
            DoEvents ' Prevent Excel from freezing.
        Next Key
        ' Move to the next row.
        rowNumber = rowNumber + 1
        DoEvents ' Prevent Excel from freezing.
    Next item

    ' Cleanup
    Set json = Nothing
    Set objRequest = Nothing
End Sub

