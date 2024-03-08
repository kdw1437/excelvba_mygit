VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VolUpdaterNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pWs As Worksheet
Private pVolatilities As Collection
Private pCodeColumn As String

' Worksheet property
Public Property Set Worksheet(ws As Worksheet)
    Set pWs = ws
End Property

Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWs
End Property

' Volatilities property
Public Property Set Volatilities(vols As Collection)
    Set pVolatilities = vols
End Property

Public Property Get Volatilities() As Collection
    Set Volatilities = pVolatilities
End Property

' CodeColumn property
Public Property Let CodeColumn(col As String)
    pCodeColumn = col
End Property

Public Property Get CodeColumn() As String
    CodeColumn = pCodeColumn
End Property

Public Sub ImportData()
    Dim volCurve As Variant
    Dim termVol As Variant
    Dim dataId As String
    Dim code As String
    Dim r As Long, c As Long
    Dim codeRow As Range
    
    For Each volCurve In pVolatilities
        dataId = volCurve("dataId")
        code = MapDataIdToCode(dataId)
        
    Next volCurve
End Sub
