VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DivStreamUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pWs As Worksheet
Private pSearchRange As Range
Private pDivStream As Collection

' Setter for the Worksheet
Public Sub SetWorksheet(sheetName As String)
    Set pWs = ThisWorkbook.Sheets(sheetName)
End Sub

' Getter for the Worksheet
Public Function GetWorksheet() As Worksheet
    Set GetWorksheet = pWs
End Function

' Setter for DivStream Collection
Public Sub SetDivStreamData(col As Collection)
    Set pDivStream = col
End Sub

' Getter for DivStream Collection
Public Function GetDivStreamData() As Collection
    Set GetDivStreamData = pDivStream
End Function

'Setter for SearchRange
Public Sub SetSearchRange(searchRange As Range)
    Set pSearchRange = searchRange
End Sub

'Getter for SearchRange
Public Function GetSearchRange() As Range
    Set GetSearchRange = pSearchRange
End Function

Public Sub UpdateWorksheet()
    If pWs Is Nothing Then
        Debug.Print "Worksheet not set."
        Exit Sub
    End If
    
    If pSearchRange Is Nothing Then
        Debug.Print "Search range not set."
        Exit Sub
    End If
    
    If pDivStream Is Nothing Then
        Debug.Print "Dividend Stream data not set."
        Exit Sub
    End If
    
    Dim i As Integer, j As Integer
    Dim dataSet As Object, divValue As Object
    Dim findCell As Range
    
    For i = 1 To pDivStream.Count
        Set dataSet = pDivStream(i)
        Set findCell = Nothing
        
        For Each findCell In pSearchRange.Cells
            If findCell.Value = dataSet("dataId") Then Exit For
            
        Next findCell
        
        If Not findCell Is Nothing Then
            For j = 1 To dataSet("dividends").Count
                Set divValue = dataSet("dividends")(j)
                findCell.Offset(j + 1, 0).Value = divValue("date")
                findCell.Offset(j + 1, 1).Value = divValue("value")
            Next j
        Else
            Debug.Print "DataId " & dataSet("dataId") & " not found in the searchRange."
        End If
    Next i
End Sub

