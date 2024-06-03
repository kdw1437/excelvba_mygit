VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "YieldCurveUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class Module: YieldCurveUpdater

Private pYieldCurves As Collection
Private pWorksheet As Worksheet
Private pCurrencyCell As Range

' YieldCurves 프로퍼티 (Collection)
Public Property Get yieldCurves() As Collection
    Set yieldCurves = pYieldCurves
End Property

Public Property Set yieldCurves(value As Collection)
    Set pYieldCurves = value
End Property

' Worksheet 프로퍼티
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWorksheet
End Property

Public Property Set Worksheet(value As Worksheet)
    Set pWorksheet = value
End Property

' CurrencyCell 프로퍼티 (Range)
Public Property Get CurrencyCell() As Range
    Set CurrencyCell = pCurrencyCell
End Property

Public Property Set CurrencyCell(value As Range)
    Set pCurrencyCell = value
End Property

' yield curve data로 worksheet를 채우는 메소드
Public Sub PopulateYieldCurveData()
    Dim yieldCurve As Variant
    Dim yieldData As Variant
    Dim FX As String
    Dim dataRow As Integer
    Dim searchRange As Range

    ' CurrencyCell property 이용
    Set searchRange = pCurrencyCell

    For Each yieldCurve In pYieldCurves
        FX = yieldCurve("dataId") ' dataId 프로퍼티의 value를 FX값으로 할당
        
        ' FX값이 적힌 셀을 프로퍼티 CurrencyCell영역에서 찾는다.
        Dim foundCurrencyCell As Range
        Set foundCurrencyCell = searchRange.Find(What:=FX, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundCurrencyCell Is Nothing Then
            ' Tenor와 Rate column을 currency cell에서 결정
            Dim tenorColumn As Integer
            tenorColumn = foundCurrencyCell.Column
            
            Dim rateColumn As Integer
            rateColumn = tenorColumn + 1 ' tenor옆에 rate가 있다.
            
            dataRow = foundCurrencyCell.row + 2
            
            ' 테이블을 정리하는 코드
            'pWorksheet.Range(pWorksheet.Cells(dataRow, tenorColumn), pWorksheet.Cells(pWorksheet.Rows.Count, rateColumn)).ClearContents
            
            For Each yieldData In yieldCurve("yields")
                'Tenor와 rate 데이터를 채워넣는다.
                pWorksheet.Cells(dataRow, tenorColumn).value = yieldData("tenor")
                pWorksheet.Cells(dataRow, rateColumn).value = yieldData("rate")
                dataRow = dataRow + 1
            Next yieldData
        End If
    Next yieldCurve
End Sub

