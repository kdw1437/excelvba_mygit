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


Public Function makeJson() As String
    Dim jsonString As String
    Dim row As Range
    Dim i As Long
    Dim jsonRow As String
    Dim exercisePrices As String
    Dim lizardBarriers As String
    Dim lizardCoupons As String
    Dim dummyCoupon As Double
    Dim earlyRedempCycle As Long
    Dim coupon As Double
    Dim adjustedRange As Range
    Set adjustedRange = pRange.Offset(1, 0).Resize(pRange.Rows.Count - 1, pRange.Columns.Count)
    
    jsonString = "["
    
    ' Skip the first row by starting with the second row (Offset(1, 0))
    For Each row In adjustedRange.Rows
        jsonRow = "{"
        
        ' Product ID, Effective Date, and Product Type
        jsonRow = jsonRow & """productId"":""" & row.Cells(1, 1).value & ""","
        jsonRow = jsonRow & """effectiveDate"":""" & row.Cells(1, 2).Text & ""","
        jsonRow = jsonRow & """productType"":""" & row.Cells(1, 3).value & ""","
        
        ' Early Redemption Cycle
        earlyRedempCycle = row.Cells(1, 4).value
        jsonRow = jsonRow & """earlyRedempCycle"":" & earlyRedempCycle & ","
        
        ' Other properties
        jsonRow = jsonRow & """settleDateOffset"":" & row.Cells(1, 5).value & ","
        jsonRow = jsonRow & """maturityEvaluationDays"":" & row.Cells(1, 6).value & ","
        jsonRow = jsonRow & """underlyingAsset1"":""" & row.Cells(1, 7).value & ""","
        jsonRow = jsonRow & """underlyingAsset2"":""" & row.Cells(1, 8).value & ""","
        jsonRow = jsonRow & """underlyingAsset3"":""" & row.Cells(1, 9).value & ""","
        
        ' Exercise Prices
        exercisePrices = ""
        For i = 10 To 15 ' Columns J to O
            If i > 10 Then exercisePrices = exercisePrices & "-"
            exercisePrices = exercisePrices & row.Cells(1, i).value
        Next i
        jsonRow = jsonRow & """exercisePrices"":""" & exercisePrices & ""","
        
        ' Coupon
        coupon = row.Cells(1, 16).value
        jsonRow = jsonRow & """coupon"":" & coupon & ","
        
        ' Dummy Coupon
        If earlyRedempCycle > 0 Then
            dummyCoupon = coupon * 3
        Else
            dummyCoupon = 0
        End If
        jsonRow = jsonRow & """dummyCoupon"":" & dummyCoupon & ","
        
        ' Remaining properties
        jsonRow = jsonRow & """lossParticipationRate"":" & row.Cells(1, 17).value & ","
        jsonRow = jsonRow & """kiBarrier"":" & row.Cells(1, 18).value & ","
        jsonRow = jsonRow & """kiBarrierCheckType"":""" & row.Cells(1, 19).value & ""","
        jsonRow = jsonRow & """koBarrier"":" & row.Cells(1, 20).value & ","
        
        ' Lizard Barriers
        lizardBarriers = row.Cells(1, 22).value & "-" & row.Cells(1, 25).value
        jsonRow = jsonRow & """lizardBarriers"":""" & lizardBarriers & ""","
        
        ' Lizard Coupons
        lizardCoupons = row.Cells(1, 23).value & "-" & row.Cells(1, 26).value
        jsonRow = jsonRow & """lizardCoupons"":""" & lizardCoupons & ""","
        
        ' Calculation Currency and Valuation
        jsonRow = jsonRow & """calculationCurrency"":""" & row.Cells(1, 27).value & ""","
        jsonRow = jsonRow & """valuation"":" & row.Cells(1, 28).value
        
        jsonRow = jsonRow & "},"
        jsonString = jsonString & jsonRow
    Next row
    
    ' Remove the last comma
    If Len(jsonString) > 1 Then
        jsonString = Left(jsonString, Len(jsonString) - 1)
    End If
    
    jsonString = jsonString & "]"
    
    makeJsonString = jsonString
End Function

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
            """exercisePrices"":" & Nz(ConstructExercisePrices(row), True) & "," & _
            """coupon"":" & Nz(row.Cells(1, 16).value, False) & "," & _
            """dummyCoupon"":" & Nz(row.Cells(1, 16).value * 3, False) & "," & _
            """lossParticipationRate"":" & Nz(row.Cells(1, 17).value, False) & "," & _
            """kiBarrier"":" & Nz(row.Cells(1, 18).value, False) & "," & _
            """kiBarrierCheckType"":" & Nz(row.Cells(1, 19).value, True) & "," & _
            """koBarrier"":" & Nz(row.Cells(1, 20).value, False) & "," & _
            """lizardBarriers"":" & Nz(ConstructLizardString(row.Cells(1, 22), row.Cells(1, 25)), True) & "," & _
            """lizardCoupons"":" & Nz(ConstructLizardString(row.Cells(1, 23), row.Cells(1, 26)), True) & "," & _
            """calculationCurrency"":" & Nz(row.Cells(1, 27).value, True) & "," & _
            """valuation"":" & Nz(row.Cells(1, 28).value, False) & _
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

' Helper function to construct exercisePrices string
Private Function ConstructExercisePrices(row As Range) As String
    Dim prices As String
    Dim i As Integer
    For i = 10 To 15
        If Not IsEmpty(row.Cells(1, i).value) Then
            If prices <> "" Then prices = prices & "-"
            prices = prices & row.Cells(1, i).value
        End If
    Next i
    If prices = "" Then prices = "null"
    ConstructExercisePrices = prices
End Function

' Helper function to construct lizardBarriers and lizardCoupons strings
Private Function ConstructLizardString(cell1 As Range, cell2 As Range) As String
    Dim part1 As String, part2 As String

    part1 = IIf(IsEmpty(cell1.value) Or cell1.value = "", "null", cell1.value)
    part2 = IIf(IsEmpty(cell2.value) Or cell2.value = "", "null", cell2.value)

    If part1 = "null" And part2 = "null" Then
        ConstructLizardString = "null"
    Else
        ConstructLizardString = part1 & "-" & part2
    End If
End Function
