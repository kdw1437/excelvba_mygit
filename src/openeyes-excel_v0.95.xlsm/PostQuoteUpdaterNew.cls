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
    Dim jsonString As String
    Dim row As Range
    Dim jsonRow As String
    Dim adjustedRange As Range
    Set adjustedRange = pRange.Offset(1, 0).Resize(pRange.Rows.Count - 1, pRange.Columns.Count)
    
    jsonString = "["
    
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
            
        jsonString = jsonString & jsonRow
    Next row
    
    jsonString = Left(jsonString, Len(jsonString) - 1) & "]"
    makeJsonString = jsonString
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

