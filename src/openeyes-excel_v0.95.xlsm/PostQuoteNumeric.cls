VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostQuoteNumeric"
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
    Dim isNumericColumn As Boolean
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
    Dim item As Variant
    Dim cellValue As Variant
    
    For rowIndex = 2 To pRange.Rows.Count
        If rowIndex > 2 Then JsonString = JsonString & ", "
        jsonRow = "{"
        
        For colIndex = 1 To pRange.Columns.Count
            ' 현재 column이 numeric column인지 확인한다.
            isNumericColumn = False
            For Each item In numericColumns
                If propertyNames.item(colIndex) = item Then
                    isNumericColumn = True
                    Exit For
                End If
            Next item
            
            cellValue = pRange.Cells(rowIndex, colIndex).value
            
            ' 빈 값을 다룬다.
            If IsEmpty(cellValue) Then
                If isNumericColumn Then
                    ' numeric column인 경우 null을 사용해서 빈 값을 표현한다.
                    jsonRow = jsonRow & """" & propertyNames.item(colIndex) & """:null"
                Else
                    ' string column인 경우 요구 사항에 따라서, empty string 혹은 null을 사용한다.
                    jsonRow = jsonRow & """" & propertyNames.item(colIndex) & """:null"
                End If
            Else
                ' numeric column이고 empty value가 아닌 경우에는 quote로 값을 감싸지 않는다.
                If isNumericColumn Then
                    jsonRow = jsonRow & """" & propertyNames.item(colIndex) & """:" & cellValue
                Else
                    jsonRow = jsonRow & """" & propertyNames.item(colIndex) & """:""" & Replace(cellValue, """", "\""") & """"
                End If
            End If
            
            If colIndex < pRange.Columns.Count Then
                jsonRow = jsonRow & ", "
            End If
        Next colIndex
        
        jsonRow = jsonRow & "}"
        
        JsonString = JsonString & jsonRow
    Next rowIndex
    
    JsonString = JsonString & "]"
    
    makeJsonString = JsonString
End Function



