Attribute VB_Name = "Module1"
Sub FetchAndWriteJSONToExcel()
    Dim httpRequest As Object
    Dim url As String
    Dim jsonResponse As String
    Dim json As Object
    Dim legsArray As Object
    Dim i As Long
    Dim j As Long
    Dim key As Variant
    Dim ws As Worksheet
    Dim cell As Range

    ' URL for the JSON data
    url = "http://localhost:8080/val/NullValuePractice"

    ' Create the XMLHTTP object and fetch the JSON data
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    With httpRequest
        .Open "GET", url, False ' Synchronous request
        .send
        jsonResponse = .responseText
    End With

    ' Parse the JSON data
    Set json = JsonConverter.ParseJson(jsonResponse)
    Set legsArray = json("response")("legsArray")

    ' Set the worksheet to write the data
    Set ws = ThisWorkbook.Sheets("Sheet1")
    ws.Cells.Clear

    ' Write the headers
    i = 1
    For Each key In legsArray(1).Keys
        ws.Cells(1, i).Value = key
        i = i + 1
    Next key

    ' Write the JSON data to the worksheet
    For i = 1 To legsArray.Count
        j = 1
        For Each key In legsArray(i).Keys
            Set cell = ws.Cells(i + 1, j)
            If IsNull(legsArray(i)(key)) Or legsArray(i)(key) = "" Then
                cell.Value = ""
            Else
                cell.Value = legsArray(i)(key)
            End If
            j = j + 1
        Next key
    Next i

    ' Format the worksheet
    ws.Columns.AutoFit

    ' Clean up
    Set httpRequest = Nothing
    Set json = Nothing
    Set legsArray = Nothing
End Sub

