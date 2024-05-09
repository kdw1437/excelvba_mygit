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
    Dim jsonRow As String
    Dim adjustedRange As Range
    Set adjustedRange = pRange.Offset(1, 0).Resize(pRange.Rows.Count - 1, pRange.Columns.Count)
    
    JsonString = "["
    
    For Each row In adjustedRange.Rows
        jsonRow = "{" & _
            """productId"":" & Nz(row.Cells(1, 1).value, False) & "," & _
            """effectiveDate"":""" & row.Cells(1, 2).Text & """," & _
            """productType"":" & Nz(row.Cells(1, 3).value, True) & "," & _
            """earlyRedempCycle"":" & Nz(row.Cells(1, 4).value, False) & "," & _
            """settleDateOffset"":" & Nz(row.Cells(1, 5).value, False) & "," & _
            """maturityEvaluationDays"":" & Nz(row.Cells(1, 6).value, False) & "," & _
            """underlyingAsset1"":" & Nz(row.Cells(1, 7).value, True) & "," & _
            """underlyingAsset2"":" & Nz(row.Cells(1, 8).value, True) & "," & _
            """underlyingAsset3"":" & Nz(row.Cells(1, 9).value, True) & "," & _
            """exercisePrices"":" & Nz(row.Cells(1, 10).value, True) & "," & _
            """coupon"":" & Nz(row.Cells(1, 11).value, False) & "," & _
            """lizardCoupon"":" & Nz(row.Cells(1, 12).value, False) & "," & _
            """lossParticipationRate"":" & Nz(row.Cells(1, 13).value, False) & "," & _
            """kiBarrier"":" & Nz(row.Cells(1, 14).value, False) & "," & _
            """principalProtectedRate"":" & Nz(row.Cells(1, 15).value, False) & "," & _
            """callBarrier"":" & Nz(row.Cells(1, 16).value, False) & "," & _
            """callParticipationRate"":" & Nz(row.Cells(1, 17).value, False) & "," & _
            """koBarrierUpSide"":" & Nz(row.Cells(1, 18).value, False) & "," & _
            """dummyCouponUpSide"":" & Nz(row.Cells(1, 19).value, False) & "," & _
            """putBarrier"":" & Nz(row.Cells(1, 20).value, False) & "," & """putParticipationRate"":" & Nz(row.Cells(1, 21).value, False) & "," & """koBarrierDownSide"":" & Nz(row.Cells(1, 22).value, False) & "," & _
            """dummyCouponDownSide"":" & Nz(row.Cells(1, 23).value, False) & "," & _
            """calculationCurrency"":" & Nz(row.Cells(1, 24).value, True) & "," & _
            """valuation"":" & Nz(row.Cells(1, 25).value, False) & _
            "},"
            
        JsonString = JsonString & jsonRow
    Next row
    
    JsonString = Left(JsonString, Len(JsonString) - 1) & "]"
    makeJsonString = JsonString
End Function

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

Private Function Nz(value As Variant, Optional isString As Boolean = False) As String
    If IsEmpty(value) Or value = "" Or value = "null" Then
        Nz = "null"
    ElseIf isString Then
        Nz = """" & value & """"
    Else
        Nz = value
    End If
End Function

' Helper function to convert empty cells to Null
Function IfEmpty(value As Variant) As Variant
    If IsEmpty(value) Or value = "" Then
        IfEmpty = Null  ' Assign Null if cell is empty
    Else
        IfEmpty = value
    End If
End Function

