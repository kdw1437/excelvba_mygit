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
' @method SetWorksheet
' @param {String} sheetName - 워크시트 이름
' @usage - 주어진 이름으로 워크시트를 설정한다.
Public Sub SetWorksheet(sheetName As String)
    Set pWs = ThisWorkbook.Sheets(sheetName)
End Sub

' Getter for the Worksheet
' @method GetWorksheet
' @return {Worksheet}
' @usage - 설정된 워크시트를 가져온다.
Public Function GetWorksheet() As Worksheet
    Set GetWorksheet = pWs
End Function

' Setter for DivData Collection
' @method SetDivData
' @param {Collection} col - DivData 컬렉션
' @usage - DivData 컬렉션을 설정한다.
Public Sub SetDivData(col As Collection)
    Set pDivData = col
End Sub

' Getter for DivData Collection
' @method GetDivData
' @return {Collection}
' @usage - 설정된 DivData 컬렉션을 가져온다.
Public Function GetDivData() As Collection
    Set GetDivData = pDivData
End Function

'Setter for DivRange
' @method SetDivRange
' @param {String} rng - 범위 주소
' @usage - DivRange 범위를 설정한다.
Public Sub SetDivRange(rng As String)
    Set pDivRange = pWs.Range(rng)
End Sub

'Getter for DivRange
' @method GetDivRange
' @return {Range}
' @usage - 설정된 DivRange 범위를 가져온다.
Public Function GetDivRange() As Range
    Set GetDivRange = pDivRange
End Function

' DivData를 사용하여 Worksheet를 업데이트하는 메서드
' @method UpdateWorksheet
' @usage - DivData 컬렉션의 데이터를 사용하여 Worksheet를 업데이트한다.
Public Sub UpdateWorksheet()

    Dim element As Variant
    Dim cell As Range
    For Each element In pDivData
        For Each cell In pDivRange
            If element("dataId") = cell.value + "_C" Then
                cell.Offset(0, 1).value = element("yield")
            End If
            DoEvents
        Next cell
        DoEvents
    Next element

End Sub

' 효율적으로 Worksheet를 업데이트하는 메서드
' @method UpdateWorksheetEfficient
' @usage - Dictionary 객체를 사용하여 Worksheet를 효율적으로 업데이트한다.
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
        DoEvents
    Next element
    
    Dim cell As Range
    
    ' 특정 Range의 cell을 iterate한다.
    For Each cell In pDivRange
        Dim cellDataId As String
        cellDataId = cell.value + "_C"
        
        ' 현재 cell의 dataId가 Dictionary에 있는지 확인한다.
        If divDataDict.Exists(cellDataId) Then
            ' 옆 칼럼에 있는 cell을 yield값으로 update한다.
            cell.Offset(0, 1).value = divDataDict(cellDataId)
        End If
        DoEvents
    Next cell

End Sub

