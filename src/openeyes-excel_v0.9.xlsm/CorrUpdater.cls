VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CorrUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pWorksheet As Worksheet
Private pSelCorrelation As Collection
Private pColumnNameRow As Integer
Private pStartRow As Integer
Private pMatrixId As String
Private pStartColumn As Integer

Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWorksheet
End Property

Public Property Set Worksheet(Value As Worksheet)
    Set pWorksheet = Value
End Property

Public Property Get SelCorrelation() As Collection
    Set SelCorrelation = pSelCorrelation
End Property

Public Property Set SelCorrelation(Value As Collection)
    Set pSelCorrelation = Value
End Property

Public Property Get ColumnNameRow() As Integer
    ColumnNameRow = pColumnNameRow
End Property

Public Property Let ColumnNameRow(Value As Integer)
    pColumnNameRow = Value
End Property

Public Property Get StartRow() As Integer
    StartRow = pStartRow
    
End Property

Public Property Let StartRow(Value As Integer)
    pStartRow = Value
End Property

Public Property Get MatrixId() As String
    MatrixId = pMatrixId
    
End Property

Public Property Let MatrixId(Value As String)
    pMatrixId = Value
End Property

Public Property Get StartColumn() As Integer
    StartColumn = pStartColumn

End Property

Public Property Let StartColumn(Value As Integer)
    pStartColumn = Value
End Property

Public Function UpdateCorrelations() As Boolean
    
    Dim LastContiguousColumn As Integer
    Dim LastContiguousRow As Integer
    Dim ColumnIndex, rowIndex As Integer
    Dim hHeader, vHeader As String

    ' 헤더에서 마지막 연속적인 column을 찾는다.
    LastContiguousColumn = pStartColumn
    While Not IsEmpty(pWorksheet.Cells(pColumnNameRow, LastContiguousColumn + 1))
        LastContiguousColumn = LastContiguousColumn + 1
    Wend
    
    ' 헤더에서 마지막 연속적인 row를 찾는다.
    LastContiguousRow = pStartRow
    While Not IsEmpty(pWorksheet.Cells(LastContiguousRow + 1, 1))
        LastContiguousRow = LastContiguousRow + 1
    Wend
    
    ' correlation으로 worksheet를 update한다.
    For ColumnIndex = pStartColumn To LastContiguousColumn
        hHeader = pWorksheet.Cells(pColumnNameRow, ColumnIndex).Value
        
        For rowIndex = pStartRow To LastContiguousRow
            vHeader = pWorksheet.Cells(rowIndex, 1).Value
            Dim correlation As Object
            For Each correlation In pSelCorrelation
                If (vHeader = correlation("dataId1") And hHeader = correlation("dataId2")) Or _
                   (vHeader = correlation("dataId2") And hHeader = correlation("dataId1")) Then
                    If correlation("matrixId") = pMatrixId Then
                        pWorksheet.Cells(rowIndex, ColumnIndex).Value = correlation("corr")
                    End If
                End If
            Next correlation
        Next rowIndex
    Next ColumnIndex
    
    UpdateCorrelations = True
End Function

