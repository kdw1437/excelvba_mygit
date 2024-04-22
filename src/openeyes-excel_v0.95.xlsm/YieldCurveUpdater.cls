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

' YieldCurves ������Ƽ (Collection)
Public Property Get YieldCurves() As Collection
    Set YieldCurves = pYieldCurves
End Property

Public Property Set YieldCurves(value As Collection)
    Set pYieldCurves = value
End Property

' Worksheet ������Ƽ
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWorksheet
End Property

Public Property Set Worksheet(value As Worksheet)
    Set pWorksheet = value
End Property

' CurrencyCell ������Ƽ (Range)
Public Property Get CurrencyCell() As Range
    Set CurrencyCell = pCurrencyCell
End Property

Public Property Set CurrencyCell(value As Range)
    Set pCurrencyCell = value
End Property

' yield curve data�� worksheet�� ä��� �޼ҵ�
Public Sub PopulateYieldCurveData()
    Dim yieldCurve As Variant
    Dim yieldData As Variant
    Dim FX As String
    Dim dataRow As Integer
    Dim searchRange As Range

    ' CurrencyCell property �̿�
    Set searchRange = pCurrencyCell

    For Each yieldCurve In pYieldCurves
        FX = yieldCurve("dataId") ' dataId ������Ƽ�� value�� FX������ �Ҵ�
        
        ' FX���� ���� ���� ������Ƽ CurrencyCell�������� ã�´�.
        Dim foundCurrencyCell As Range
        Set foundCurrencyCell = searchRange.Find(What:=FX, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundCurrencyCell Is Nothing Then
            ' Tenor�� Rate column�� currency cell���� ����
            Dim tenorColumn As Integer
            tenorColumn = foundCurrencyCell.Column
            
            Dim rateColumn As Integer
            rateColumn = tenorColumn + 1 ' tenor���� rate�� �ִ�.
            
            dataRow = foundCurrencyCell.row + 2
            
            ' ���̺��� �����ϴ� �ڵ�
            'pWorksheet.Range(pWorksheet.Cells(dataRow, tenorColumn), pWorksheet.Cells(pWorksheet.Rows.Count, rateColumn)).ClearContents
            
            For Each yieldData In yieldCurve("yields")
                'Tenor�� rate �����͸� ä���ִ´�.
                pWorksheet.Cells(dataRow, tenorColumn).value = yieldData("tenor")
                pWorksheet.Cells(dataRow, rateColumn).value = yieldData("rate")
                dataRow = dataRow + 1
            Next yieldData
        End If
    Next yieldCurve
End Sub
