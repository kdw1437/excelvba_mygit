VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostQuoteUpdaterNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pWs As Worksheet
Private pRange As Range

' Getter for pWs
' @method Worksheet
' @return {Worksheet}
' @usage - worksheet을 get한다.
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWs
End Property

' Setter for pWs
' @method Worksheet
' @param {Worksheet} value
' @usage - worksheet을 set한다.
Public Property Set Worksheet(value As Worksheet)
    Set pWs = value
End Property

' Getter for pRange
' @method Range
' @return {Range}
' @usage - range를 get한다.
Public Property Get Range() As Range
    Set Range = pRange
End Property

' Setter for pRange
' @method Range
' @param {Range} value
' @usage - range를 set한다.
Public Property Set Range(value As Range)
    Set pRange = value
End Property


' 특정 영역에 있는 data로 부터 JSON string을 만든다.
' @method makeJsonString2
' @return {String}
' @usage - 특정 영역에 있는 data로 부터 JSON string을 만든다.
Public Function makeJsonString2() As String

    Dim JsonString As String
    Dim row As Range
    Dim adjustedRange As Range
    Set adjustedRange = pRange.Offset(1, 0).Resize(pRange.Rows.Count - 1, pRange.Columns.Count)
    Dim JsonCollection As Collection
    Set JsonCollection = New Collection
    
    For Each row In adjustedRange.Rows
        Dim JsonDictionary As Dictionary
        Set JsonDictionary = New Dictionary
        
        JsonDictionary.Add "productId", row.Cells(1, 1).value
        JsonDictionary.Add "effectiveDate", IfEmpty(CStr(row.Cells(1, 2).value))
        JsonDictionary.Add "productType", row.Cells(1, 3).value
        JsonDictionary.Add "earlyRedempCycle", row.Cells(1, 4).value
        JsonDictionary.Add "settleDateOffset", row.Cells(1, 5).value
        JsonDictionary.Add "maturityEvaluationDays", row.Cells(1, 6).value
        JsonDictionary.Add "underlyingAsset1", IfEmpty(row.Cells(1, 7).value)
        JsonDictionary.Add "underlyingAsset2", IfEmpty(row.Cells(1, 8).value)
        JsonDictionary.Add "underlyingAsset3", IfEmpty(row.Cells(1, 9).value)
        JsonDictionary.Add "exercisePrices", IfEmpty(CStr(row.Cells(1, 10).value))
        JsonDictionary.Add "coupon", IfEmpty(row.Cells(1, 11).value)
        JsonDictionary.Add "lizardCoupon", IfEmpty(row.Cells(1, 12).value)
        JsonDictionary.Add "lossParticipationRate", IfEmpty(row.Cells(1, 13).value)
        JsonDictionary.Add "kiBarrier", IfEmpty(row.Cells(1, 14).value)
        JsonDictionary.Add "principalProtectedRate", IfEmpty(row.Cells(1, 15).value)
        JsonDictionary.Add "callBarrier", IfEmpty(row.Cells(1, 16).value)
        JsonDictionary.Add "callParticipationRate", IfEmpty(row.Cells(1, 17).value)
        JsonDictionary.Add "koBarrierUpSide", IfEmpty(row.Cells(1, 18).value)
        JsonDictionary.Add "dummyCouponUpSide", IfEmpty(row.Cells(1, 19).value)
        JsonDictionary.Add "putBarrier", IfEmpty(row.Cells(1, 20).value)
        JsonDictionary.Add "putParticipationRate", IfEmpty(row.Cells(1, 21).value)
        JsonDictionary.Add "koBarrierDownSide", IfEmpty(row.Cells(1, 22).value)
        JsonDictionary.Add "dummyCouponDownSide", IfEmpty(row.Cells(1, 23).value)
        JsonDictionary.Add "calculationCurrency", IfEmpty(row.Cells(1, 24).value)
        JsonDictionary.Add "valuation", IfEmpty(row.Cells(1, 25).value)
        
        JsonCollection.Add JsonDictionary
        
    Next row
    
    JsonString = JsonConverter.ConvertToJson(JsonCollection)
    makeJsonString2 = JsonString
    
    
End Function

' empty cell을 Null로 변경하는 Helper 함수
' @method IfEmpty
' @param {Variant} value
' @return {Variant}
' @usage - empty cell을 Null로 변경, 다른 경우에는 cell값을 return.
Function IfEmpty(value As Variant) As Variant
    If IsEmpty(value) Or value = "" Then
        IfEmpty = Null  ' 빈 cell인 경우, Null 할당.
    Else
        IfEmpty = value
    End If
End Function

