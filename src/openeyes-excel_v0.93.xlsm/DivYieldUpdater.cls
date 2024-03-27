VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DivYieldUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pWs As Worksheet
Private pDivRange As Range
Private pDivData As Collection

' Setter for the Worksheet
Public Sub SetWorksheet(sheetName As String)
    Set pWs = ThisWorkbook.Sheets(sheetName)
End Sub

' Getter for the Worksheet
Public Function GetWorksheet() As Worksheet
    Set GetWorksheet = pWs
End Function

' Setter for DivData Collection
Public Sub SetDivData(col As Collection)
    Set pDivData = col
End Sub

' Getter for DivData Collection
Public Function GetDivData() As Collection
    Set GetDivData = pDivData
End Function

'Setter for DivRange
Public Sub SetDivRange(rng As String)
    Set pDivRange = pWs.Range(rng)
End Sub

'Getter for DivRange
Public Function GetDivRange() As Range
    Set GetDivRange = pDivRange
End Function

Public Sub UpdateWorksheet()

    Dim element As Variant
    Dim cell As Range
    For Each element In pDivData
        For Each cell In pDivRange
            If element("dataId") = cell.Value + "_C" Then
                cell.Offset(0, 1).Value = element("yield")
            End If
        Next cell
    Next element

End Sub
