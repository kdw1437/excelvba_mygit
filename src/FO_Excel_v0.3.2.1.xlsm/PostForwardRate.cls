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

Private pWorksheet As Worksheet
Private pStartCell As Range
Private pK As Long

' Property for the worksheet
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWorksheet
End Property

Public Property Set Worksheet(value As Worksheet)
    Set pWorksheet = value
End Property

' Property for the start cell
Public Property Get StartCell() As Range
    Set StartCell = pStartCell
End Property

Public Property Set StartCell(value As Range)
    Set pStartCell = value
End Property

' Property for the number of currencies
Public Property Get K() As Long
    K = pK
End Property

Public Property Let K(value As Long)
    pK = value
End Property

Public Sub GenerateJSON()
    Dim currenciesCollection As New Collection
    Dim i As Long, j As Long
    
    For i = 1 To pK ' Use the K property here
        Dim relCurrencyCell As Range
        Dim currencyCode As String
        currencyCode = pStartCell.Offset(4, 1 + 3 * (i - 1)).value
        
        Set relCurrencyCell = pWorksheet.Rows(7).Find(What:=currencyCode, Lookat:=xlWhole)
        
        Dim currencyData As Object
        Set currencyData = CreateObject("Scripting.Dictionary")
        
        Dim dataId As String
        dataId = "F_FX_" & currencyCode & pStartCell.Offset(3, 1 + 3 * (i - 1)).value
        currencyData("dataId") = dataId
        
        Dim yieldsCollection As New Collection
        Set yieldsCollection = New Collection '재초기화
        
        j = 2 ' Start row offset for yields
        Do While pWorksheet.Cells(relCurrencyCell.Row + j, relCurrencyCell.Column).value <> ""
            Dim yieldData As Object
            Set yieldData = CreateObject("Scripting.Dictionary")
            yieldData("tenor") = pWorksheet.Cells(relCurrencyCell.Row + j, relCurrencyCell.Column - 1).value
            yieldData("value") = pWorksheet.Cells(relCurrencyCell.Row + j, relCurrencyCell.Column).value
            yieldsCollection.Add yieldData
            j = j + 1
        Loop
        
        Dim yieldsArray() As Variant
        Dim idx As Long
        
        ReDim yieldsArray(1 To yieldsCollection.Count)
        For idx = 1 To yieldsCollection.Count
            Set yieldsArray(idx) = yieldsCollection(idx)
        Next idx
        
        currencyData("yields") = yieldsArray
        currenciesCollection.Add currencyData
    Next i
    
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(currenciesCollection)
    
    Debug.Print jsonString
End Sub

