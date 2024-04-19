VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostForwardRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pArrayOfDataId As Variant
Private pYieldCurveRow As Range
Private pStartCell As Range

' Getter for pArrayOfDataId
Public Property Get ArrayOfDataId() As Variant
    ArrayOfDataId = pArrayOfDataId
End Property

' Setter for pArrayOfDataId
Public Property Let ArrayOfDataId(ByVal value As Variant)
    pArrayOfDataId = value
End Property

' Getter for pBaseCurrencyRow
Public Property Get BaseCurrencyRow() As Range
    Set BaseCurrencyRow = pBaseCurrencyRow
End Property

' Setter for pBaseCurrencyRow
Public Property Set BaseCurrencyRow(ByVal value As Range)
    Set pBaseCurrencyRow = value
End Property


' Setter for pStartCell
Public Property Set startCell(value As Range)
    Set pStartCell = value
End Property

' Getter for pStartCell
Public Property Get startCell() As Range
    Set startCell = pStartCell
End Property

Sub makeJsonString()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Missing Data - Fx Forward")
    
    ws.pStartCell
End Sub
