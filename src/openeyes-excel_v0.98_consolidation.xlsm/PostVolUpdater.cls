VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostVolUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private 변수
Private pWorksheet As Worksheet
Private pRefCell As Range
Private pDataId As String

' worksheet property
' @method Worksheet
' @return {Worksheet}
' @usage - worksheet의 getter
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWorksheet
End Property

' @method Worksheet
' @param {Worksheet} value
' @usage - worksheet의 setter
Public Property Set Worksheet(value As Worksheet)
    Set pWorksheet = value
End Property

' reference cell property
' @method RefCell
' @return {Range}
' @usage - reference cell의 getter
Public Property Get RefCell() As Range
    Set RefCell = pRefCell
End Property

' @method RefCell
' @param {Range} value
' @usage - reference cell의 setter
Public Property Set RefCell(value As Range)
    Set pRefCell = value
End Property

' dataId property
' @method dataId
' @return {String}
' @usage - dataId의 getter
Public Property Get dataId() As String
    dataId = pDataId
End Property

' @method dataId
' @param {String} value
' @usage - dataId의 setter
Public Property Let dataId(value As String)
    pDataId = value
End Property

' GenerateObjectJSON method
' @method GenerateObjectJSON
' @return {String}
' @usage - worksheet의 tenor, vol, volFactor data를 바탕으로 JSON string을 생성
Public Function GenerateObjectJSON() As String
    Dim volFactorRange As Range, tenorRange As Range, dataRange As Range
    Dim volFactorCell As Range, termVolCell As Range
    Dim firstTermVol As Boolean, firstVolCurve As Boolean
    Dim objectJSON As String
    
    ' refCell의 데이터와 volFactor, tenor에 근거해서 range를 잡는다.
    Set volFactorRange = pWorksheet.Range(pRefCell.Offset(0, 2), pRefCell.Offset(0, 2).End(xlToRight))
    Set tenorRange = pWorksheet.Range(pRefCell.Offset(1, 1), pRefCell.Offset(1, 1).End(xlDown))
    Set dataRange = pWorksheet.Range(volFactorRange.Offset(1, 0), tenorRange.Offset(0, volFactorRange.Columns.Count - 1))
    
    objectJSON = "{" & """dataId"": """ & pDataId & """," & """volCurves"": ["
    
    firstVolCurve = True
    For Each volFactorCell In volFactorRange
        If Not firstVolCurve Then
            objectJSON = objectJSON & ","
        End If
        objectJSON = objectJSON & "{" & """termVols"": ["
        
        firstTermVol = True
        For Each termVolCell In tenorRange
            If Not firstTermVol Then
                objectJSON = objectJSON & ","
            End If
            objectJSON = objectJSON & "{" & """tenor"": " & termVolCell.value & "," & """vol"": " & pWorksheet.Cells(termVolCell.row, volFactorCell.Column).value & "}"
            firstTermVol = False
        Next termVolCell
        
        objectJSON = objectJSON & "]," & """volFactor"": " & volFactorCell.value & "}"
        firstVolCurve = False
    Next volFactorCell
    
    objectJSON = objectJSON & "]}"
    
    GenerateObjectJSON = objectJSON
End Function

' GenerateObjectJSON2 method
' @method GenerateObjectJSON
' @return {String}
' @usage - worksheet의 tenor, vol, volFactor data를 바탕으로 JSON string을 생성
Public Function GenerateObjectJSON2() As String
    Dim volFactorRange As Range, tenorRange As Range, dataRange As Range
    Dim volFactorCell As Range, termVolCell As Range
    Dim volCurves As Collection
    Dim termVols As Collection
    Dim volCurve As Dictionary, termVol As Dictionary
    Dim volFactorOrder As Integer
    
    ' Initialize the main collection for volCurves
    Set volCurves = New Collection
    
    ' Define the ranges based on the refCell
    Set volFactorRange = pWorksheet.Range(pRefCell.Offset(0, 2), pRefCell.Offset(0, 2).End(xlToRight))
    Set tenorRange = pWorksheet.Range(pRefCell.Offset(1, 1), pRefCell.Offset(1, 1).End(xlDown))
    Set dataRange = pWorksheet.Range(volFactorRange.Offset(1, 0), tenorRange.Offset(0, volFactorRange.Columns.Count - 1))
    
    volFactorOrder = 1
    ' Iterate over each volFactor
    For Each volFactorCell In volFactorRange
        Set termVols = New Collection
        ' Create termVols for each tenor
        For Each termVolCell In tenorRange
            Set termVol = New Dictionary
            termVol.Add "tenor", termVolCell.value
            termVol.Add "vol", pWorksheet.Cells(termVolCell.row, volFactorCell.Column).value
            termVols.Add termVol
            DoEvents
        Next termVolCell
        
        Set volCurve = New Dictionary
        volCurve.Add "volFactor", volFactorCell.value
        volCurve.Add "volFactorOrder", volFactorOrder
        volCurve.Add "termVols", termVols
        volCurves.Add volCurve
        DoEvents
        
        volFactorOrder = volFactorOrder + 1
    Next volFactorCell
    
    ' Create the main dictionary to hold dataId and volCurves
    Dim output As Dictionary
    Set output = New Dictionary
    output.Add "dataId", pDataId
    output.Add "volCurves", volCurves
    
    ' Convert to JSON using jsonConverter
    GenerateObjectJSON2 = JsonConverter.ConvertToJson(output)
End Function

