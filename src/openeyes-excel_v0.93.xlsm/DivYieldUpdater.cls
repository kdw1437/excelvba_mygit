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

Public Sub UpdateWorksheetEfficient()

    ' dataIds의 yields에 대한 mapping을 가지는 Dictionary 객체를 만든다.
    Dim divDataDict As Object
    Set divDataDict = CreateObject("Scripting.Dictionary")
    
    Dim element As Variant
    
    ' dataIds값을 key로 yields값을 item으로 Dictionary 객체를 채운다.
    For Each element In pDivData
        Dim dataId As String
        dataId = element("dataId")
        Dim yield As Variant
        yield = element("yield")
        divDataDict(dataId) = yield
    Next element
    
    Dim cell As Range
    
    ' 특정 Range의 cell을 iterate한다.
    For Each cell In pDivRange
        Dim cellDataId As String
        cellDataId = cell.Value + "_C"
        
        ' 현재 cell의 dataId가 Dictionary에 있는지 확인한다.
        If divDataDict.Exists(cellDataId) Then
            ' 옆 칼럼에 있는 cell을 yield값으로 update한다.
            cell.Offset(0, 1).Value = divDataDict(cellDataId)
        End If
    Next cell

End Sub

