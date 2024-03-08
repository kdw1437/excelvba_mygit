VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VolUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pWs As Worksheet
Private pSearchRange1 As Range
Private pSearchRange2 As Range

Public Property Set Worksheet(ws As Worksheet)
    Set pWs = ws
End Property

Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWs
End Property

Public Property Set Range1(rng As Range)
    Set pSearchRange1 = rng
End Property

Public Property Get Range1() As Range
    Set Range1 = pSearchRange1
End Property

Public Property Set Range2(rng As Range)
    Set pSearchRange2 = rng
End Property

Public Property Get Range2() As Range
    Set Range2 = pSearchRange2
End Property

Public Sub ProcessVolCurve(volCurves As Collection)
    Dim volCurve As Variant
    Dim termVol As Variant
    Dim dataId As String
    Dim code As String
    Dim r As Long, c As Long
    
    For Each volCurve In volCurves
        dataId = volCurve("dataId")
        code = MapDataIdToCode(dataId)
        
        
    Next volCurve
End Sub
