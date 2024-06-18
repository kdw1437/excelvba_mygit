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

' Worksheet property
' @method Worksheet
' @return {Worksheet}
' @usage - pWorksheet 값을 가져온다.
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWorksheet
End Property

' @method Worksheet
' @param {Worksheet} value
' @usage - pWorksheet 값을 설정한다.
Public Property Set Worksheet(value As Worksheet)
    Set pWorksheet = value
End Property

' SelCorrelation property
' @method SelCorrelation
' @return {Collection}
' @usage - pSelCorrelation 값을 가져온다.
Public Property Get SelCorrelation() As Collection
    Set SelCorrelation = pSelCorrelation
End Property

' @method SelCorrelation
' @param {Collection} value
' @usage - pSelCorrelation 값을 설정한다.
Public Property Set SelCorrelation(value As Collection)
    Set pSelCorrelation = value
End Property

' ColumnNameRow property
' @method ColumnNameRow
' @return {Integer}
' @usage - pColumnNameRow 값을 가져온다.
Public Property Get ColumnNameRow() As Integer
    ColumnNameRow = pColumnNameRow
End Property

' @method ColumnNameRow
' @param {Integer} value
' @usage - pColumnNameRow 값을 설정한다.
Public Property Let ColumnNameRow(value As Integer)
    pColumnNameRow = value
End Property

' StartRow property
' @method StartRow
' @return {Integer}
' @usage - pStartRow 값을 가져온다.
Public Property Get startRow() As Integer
    startRow = pStartRow
End Property

' @method StartRow
' @param {Integer} value
' @usage - pStartRow 값을 설정한다.
Public Property Let startRow(value As Integer)
    pStartRow = value
End Property

' MatrixId property
' @method MatrixId
' @return {String}
' @usage - pMatrixId 값을 가져온다.
Public Property Get MatrixId() As String
    MatrixId = pMatrixId
End Property

' @method MatrixId
' @param {String} value
' @usage - pMatrixId 값을 설정한다.
Public Property Let MatrixId(value As String)
    pMatrixId = value
End Property

' StartColumn property
' @method StartColumn
' @return {Integer}
' @usage - pStartColumn 값을 가져온다.
Public Property Get StartColumn() As Integer
    StartColumn = pStartColumn
End Property

' @method StartColumn
' @param {Integer} value
' @usage - pStartColumn 값을 설정한다.
Public Property Let StartColumn(value As Integer)
    pStartColumn = value
End Property

' correlation 데이터를 worksheet에 업데이트하는 함수
' @method UpdateCorrelations
' @return {Boolean}
' @usage - SelCorrelation 컬렉션의 데이터를 기반으로 worksheet의 correlation 값을 업데이트한다.
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
        hHeader = pWorksheet.Cells(pColumnNameRow, ColumnIndex).value
        
        For rowIndex = pStartRow To LastContiguousRow
            vHeader = pWorksheet.Cells(rowIndex, 1).value
            Dim correlation As Object
            For Each correlation In pSelCorrelation
                If (vHeader = correlation("dataId1") And hHeader = correlation("dataId2")) Or _
                   (vHeader = correlation("dataId2") And hHeader = correlation("dataId1")) Then
                    If correlation("matrixId") = pMatrixId Then
                        pWorksheet.Cells(rowIndex, ColumnIndex).value = correlation("corr")
                    End If
                End If
            Next correlation
            
            DoEvents
            
        Next rowIndex
        
        DoEvents
        
    Next ColumnIndex
    
    UpdateCorrelations = True
End Function

