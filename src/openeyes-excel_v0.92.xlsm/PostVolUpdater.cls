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
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWorksheet
End Property

Public Property Set Worksheet(Value As Worksheet)
    Set pWorksheet = Value
End Property

' reference cell property
Public Property Get RefCell() As Range
    Set RefCell = pRefCell
End Property

Public Property Set RefCell(Value As Range)
    Set pRefCell = Value
End Property

' dataId property
Public Property Get dataId() As String
    dataId = pDataId
End Property

Public Property Let dataId(Value As String)
    pDataId = Value
End Property

' GenerateObjectJSON method
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
            objectJSON = objectJSON & "{" & """tenor"": " & termVolCell.Value & "," & """vol"": " & pWorksheet.Cells(termVolCell.Row, volFactorCell.Column).Value & "}"
            firstTermVol = False
        Next termVolCell
        
        objectJSON = objectJSON & "]," & """volFactor"": " & volFactorCell.Value & "}"
        firstVolCurve = False
    Next volFactorCell
    
    objectJSON = objectJSON & "]}"
    
    GenerateObjectJSON = objectJSON
End Function

