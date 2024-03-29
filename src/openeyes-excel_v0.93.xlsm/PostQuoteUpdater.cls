VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostQuoteUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pWs As Worksheet
Private pRange As Range

'Getter for pWs
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWs
End Property

'Setter for pWs
Public Property Set Worksheet(value As Worksheet)
    Set pWs = value
End Property

'Getter for pRange
Public Property Get Range() As Range
    Set Range = pRange
End Property

'Setter for pRange
Public Property Set Range(value As Range)
    Set pRange = value
End Property

Public Function makeJsonString() As String

    Dim jsonString As String
    Dim row As Range
    Dim cell As Range
    Dim i As Long
    Dim jsonRow As String
    Dim propertyNames As Collection
    Dim rowIndex As Long
    Dim colIndex As Long
    
    Set propertyNames = New Collection
    
    propertyNames.Add "SEQ"
    propertyNames.Add "transactionDate"
    propertyNames.Add "productStructureDetails"
    propertyNames.Add "earlyRepaymentCycle"
    propertyNames.Add "settlementDateOffset"
    propertyNames.Add "maturityEvaluationDays"
    propertyNames.Add "underlyingAsset1"
    propertyNames.Add "underlyingAsset2"
    propertyNames.Add "underlyingAsset3"
    propertyNames.Add "exercisePrice1"
    propertyNames.Add "exercisePrice2"
    propertyNames.Add "exercisePrice3"
    propertyNames.Add "exercisePrice4"
    propertyNames.Add "exercisePrice5"
    propertyNames.Add "exercisePrice6"
    propertyNames.Add "coupon"
    propertyNames.Add "lossParticipationRate"
    propertyNames.Add "KIBarrier"
    propertyNames.Add "KOBarrier"
    propertyNames.Add "lizardFirstOpportunity"
    propertyNames.Add "lizardFirstBarrier"
    propertyNames.Add "lizardFirstCoupon"
    propertyNames.Add "lizardSecondOpportunity"
    propertyNames.Add "lizardSecondBarrier"
    propertyNames.Add "lizardSecondCoupon"
    propertyNames.Add "calculationCurrency"
    propertyNames.Add "valuation"
    
    jsonString = "["
    
    For rowIndex = 2 To pRange.Rows.Count
        If rowIndex > 2 Then jsonString = jsonString & ", "
        jsonRow = "{"
        For colIndex = 1 To pRange.Columns.Count
            jsonRow = jsonRow & """" & propertyNames.Item(colIndex) & """:""" & Replace(pRange.Cells(rowIndex, colIndex).value, """", "\""") & """"
            If colIndex < pRange.Columns.Count Then
                jsonRow = jsonRow & ", "
            End If
        Next colIndex
        
        jsonRow = jsonRow & "}"
        
'        If rowIndex < pRange.Rows.Count Then
'            jsonRow = jsonRow & ", "
'        End If
        
        jsonString = jsonString & jsonRow
    Next rowIndex
    
    jsonString = jsonString & "]"
    
    makeJsonString = jsonString
End Function


