VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostDivYieldUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pWs As Worksheet
Private pDivCell As Range
Private pdataIdRange As Range

' Getter for Worksheet
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWs
End Property

' Setter for Worksheet
Public Property Set Worksheet(value As Worksheet)
    Set pWs = value
End Property

' Getter for DivCell Range
Public Property Get DivCell() As Range
    Set DivCell = pDivCell
End Property

' Setter for DivCell Range
Public Property Set DivCell(value As Range)
    Set pDivCell = value
End Property

' Getter for DataIdRange
Public Property Get DataIdRange() As Range
    Set DataIdRange = pdataIdRange
End Property

' Setter for DataIdRange
Public Property Set DataIdRange(value As Range)
    Set pdataIdRange = value
End Property

' Setter for DivCell Range using a string to define the range
'Public Property Let DivCell(Address As String)
'    Set pDivCell = pWs.Range(Address)
'End Property

' Setter for DataIdRange using a string to define the range
'Public Property Let DataIdRange(Address As String)
'    Set pdataIdRange = pWs.Range(Address)
'End Property


Public Function GenerateJson() As String
    Dim jsonString As String
    jsonString = "["
    Dim cell As Range
    Dim dataId As String
    Dim yield As Double
    
    For Each cell In pdataIdRange
        dataId = cell.value
        yield = cell.Offset(0, 1).value
        Dim jsonObject As String
        jsonObject = "{""dataId"": """ & dataId & """, ""yield"": " & yield & "}"
        
        If Len(jsonString) > 1 Then
            jsonString = jsonString & ", "
        End If
        
        jsonString = jsonString & jsonObject
        
    Next cell
    
    jsonString = jsonString & "]"
    
    GenerateJson = jsonString
End Function
