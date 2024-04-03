VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostPriceUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pTable1Point As Range
Private pFxRow As Range
Private pWs As Worksheet

Public Property Set Table1Point(value As Range)
    Set pTable1Point = value
End Property

Public Property Get Table1Point() As Range
    Set Table1Point = pTable1Point
End Property

Public Property Set fxRow(value As Range)
    Set pFxRow = value
End Property

Public Property Get fxRow() As Range
    Set fxRow = pFxRow
End Property

Public Property Set Worksheet(value As Worksheet)
    Set pWs = value
End Property

Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWs
End Property

' Method to construct JSON string
Public Function GenerateJson() As String
    Dim DataString As String
    DataString = "[" ' Initialize the DataString as an empty JSON array
    
    Dim i As Integer
    For i = pTable1Point.row + 1 To pFxRow.row - 2
        Dim dataId As String
        Dim closePric As String
        
        dataId = pWs.Cells(i, pTable1Point.Column).value
        closePric = pWs.Cells(i, pTable1Point.Column + 1).value
        
        ' Construct the JSON object for the current row
        Dim jsonObject As String
        jsonObject = "{""dataId"": """ & dataId & """, ""price"": " & closePric & "}"
        
        ' If not the first item, add a comma separator
        If Len(DataString) > 1 Then
            DataString = DataString & ", "
        End If
        
        ' Append the JSON object to the DataString (JSON array)
        DataString = DataString & jsonObject
    Next i

    DataString = DataString & "]" ' Close the JSON array
    
    GenerateJson = DataString
End Function

