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

    Dim JsonString As String
    Dim row As Range
    Dim cell As Range
    Dim i As Long
    Dim jsonRow As String
    Dim propertyNames As Collection
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim numericColumns As Collection
    
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
    
    'value가 numeric인 column의 경우, 새로운 collection을 만들어서 이 column의 값들을 numeric으로 처리해준다.
    Set numericColumns = New Collection
    numericColumns.Add "SEQ"
    numericColumns.Add "earlyRepaymentCycle"
    numericColumns.Add "settlementDateOffset"
    numericColumns.Add "maturityEvaluationDays"
    numericColumns.Add "exercisePrice1"
    numericColumns.Add "exercisePrice2"
    numericColumns.Add "exercisePrice3"
    numericColumns.Add "exercisePrice4"
    numericColumns.Add "exercisePrice5"
    numericColumns.Add "exercisePrice6"
    numericColumns.Add "coupon"
    numericColumns.Add "lossParticipationRate"
    numericColumns.Add "KIBarrier"
    numericColumns.Add "KOBarrier"
    numericColumns.Add "lizardFirstOpportunity"
    numericColumns.Add "lizardFirstBarrier"
    numericColumns.Add "lizardFirstCoupon"
    numericColumns.Add "lizardSecondOpportunity"
    numericColumns.Add "lizardSecondBarrier"
    numericColumns.Add "lizardSecondCoupon"
    numericColumns.Add "valuation"
    
    
    JsonString = "["
    
    For rowIndex = 2 To pRange.Rows.Count
        If rowIndex > 2 Then JsonString = JsonString & ", "
        jsonRow = "{"
        For colIndex = 1 To pRange.Columns.Count
            jsonRow = jsonRow & """" & propertyNames.item(colIndex) & """:""" & Replace(pRange.Cells(rowIndex, colIndex).value, """", "\""") & """"
            If colIndex < pRange.Columns.Count Then
                jsonRow = jsonRow & ", "
            End If
        Next colIndex
        
        jsonRow = jsonRow & "}"
        
'        If rowIndex < pRange.Rows.Count Then
'            jsonRow = jsonRow & ", "
'        End If
        
        JsonString = JsonString & jsonRow
    Next rowIndex
    
    JsonString = JsonString & "]"
    
    makeJsonString = JsonString
End Function


