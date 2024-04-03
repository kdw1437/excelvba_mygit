VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostQuoteRecent"
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
    Dim isNumericColumn As Boolean
    Dim numericColumns As Collection
    
    Set propertyNames = New Collection
    
    propertyNames.Add "productId"
    propertyNames.Add "effectiveDate"
    propertyNames.Add "productType"
    propertyNames.Add "earlyRedempCycle"
    propertyNames.Add "settleDateOffset"
    propertyNames.Add "maturityEvaluationDays"
    propertyNames.Add "underlyingAsset1"
    propertyNames.Add "underlyingAsset2"
    propertyNames.Add "underlyingAsset3"
    propertyNames.Add "exercisePrices"
    propertyNames.Add "coupon"
    propertyNames.Add "dummyCoupon"
    propertyNames.Add "lossParticipationRate"
    propertyNames.Add "kiBarrier"
    propertyNames.Add "kiBarrierCheckType"
    propertyNames.Add "koBarrier"
    propertyNames.Add "lizardBarriers"
    propertyNames.Add "lizardCoupons"
    propertyNames.Add "calculationCurrency"
    propertyNames.Add "valuation"
    
    'value가 numeric인 column의 경우, 새로운 collection을 만들어서 이 column의 값들을 numeric으로 처리해준다.
    Set numericColumns = New Collection
    numericColumns.Add "earlyRedempCycle"
    numericColumns.Add "settleDateOffset"
    numericColumns.Add "maturityEvaluationDays"
    numericColumns.Add "coupon"
    numericColumns.Add "dummyCoupon"
    numericColumns.Add "lossParticipationRate"
    numericColumns.Add "kiBarrier"
    numericColumns.Add "koBarrier"
    numericColumns.Add "valuation"


