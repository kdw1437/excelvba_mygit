VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VolUpdaterNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pWs As Worksheet
Private pVolatilities As Collection
Private pCodeColumn As String

' Worksheet property
' @method Worksheet
' @param {Worksheet} ws
' @usage - pWs ���� �����Ѵ�.
Public Property Set Worksheet(ws As Worksheet)
    Set pWs = ws
End Property

' @method Worksheet
' @return {Worksheet}
' @usage - pWs ���� �����´�.
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWs
End Property

' Volatilities property
' @method Volatilities
' @param {Collection} vols
' @usage - pVolatilities ���� �����Ѵ�.
Public Property Set Volatilities(vols As Collection)
    Set pVolatilities = vols
End Property

' @method Volatilities
' @return {Collection}
' @usage - pVolatilities ���� �����´�.
Public Property Get Volatilities() As Collection
    Set Volatilities = pVolatilities
End Property

' CodeColumn property
' @method CodeColumn
' @param {String} col
' @usage - pCodeColumn ���� �����Ѵ�.
Public Property Let CodeColumn(col As String)
    pCodeColumn = col
End Property

' @method CodeColumn
' @return {String}
' @usage - pCodeColumn ���� �����´�.
Public Property Get CodeColumn() As String
    CodeColumn = pCodeColumn
End Property

' �����͸� �������� �޼���
' @method ImportData
' @usage - pVolatilities �÷����� �����͸� Excel ��ũ��Ʈ�� �����´�.
Public Sub ImportData()
    Dim volCurve As Variant
    Dim termVol As Variant
    Dim dataId As String
    Dim code As String
    Dim r As Long, c As Long
    Dim codeRow As Range
    Dim volEntry As Variant
    
    
    For Each volCurve In pVolatilities
        dataId = volCurve("dataId")
        code = MapDataIdToCode(dataId)
        
        Set codeRow = pWs.Columns(pCodeColumn).Find(What:=code, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not codeRow Is Nothing Then
            Dim codeRowNumber As Long
            codeRowNumber = codeRow.row
            Dim searchRange1 As Range
            Dim searchRange2 As Range
            Set searchRange1 = pWs.Range(pWs.Cells(codeRowNumber, 3), pWs.Cells(codeRowNumber, 3).End(xlToRight))
            Set searchRange2 = pWs.Range(pWs.Cells(codeRowNumber + 1, 2), pWs.Cells(codeRowNumber + 1, 2).End(xlDown))
            
            For Each termVol In volCurve("volCurves")
                Dim volFactor As Double
                volFactor = termVol("volFactor")
                
                c = pWs.Rows(codeRowNumber).Find(What:=volFactor, LookIn:=xlValues, LookAt:=xlWhole).Column
                For Each volEntry In termVol("termVols")
                    Dim tenor As Double
                    tenor = volEntry("tenor")
                    
                    Dim tenorCell As Range
                    Set tenorCell = searchRange2.Find(What:=tenor, LookIn:=xlValues, LookAt:=xlWhole)
                    
                    If Not tenorCell Is Nothing Then
                        r = tenorCell.row
                        pWs.Cells(r, c).value = volEntry("vol")
                    End If
                    DoEvents
                Next volEntry
                DoEvents
            Next termVol
        End If
        DoEvents
    Next volCurve
End Sub

' �� ���� ä��� �޼���
' @method FillEmptyCells
' @usage - pVolatilities �÷����� �����͸� ������� Excel ��ũ��Ʈ���� �� ���� 0���� ä���.
Public Sub FillEmptyCells()
    Dim headerCell As Range
    Dim rowHeaderCell As Range
    Dim dataCell As Range
    Dim searchRange1 As Range, searchRange2 As Range
    Dim code As String
    Dim codeRow As Range
    Dim volCurve As Variant
    Dim dataId As String
            
    For Each volCurve In pVolatilities
        dataId = volCurve("dataId")
        code = MapDataIdToCode(dataId)
        
        Set codeRow = pWs.Columns(pCodeColumn).Find(What:=code, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not codeRow Is Nothing Then
            Dim codeRowNumber As Long
            codeRowNumber = codeRow.row
           
            Set searchRange1 = pWs.Range(pWs.Cells(codeRowNumber, 3), pWs.Cells(codeRowNumber, 3).End(xlToRight))
            Set searchRange2 = pWs.Range(pWs.Cells(codeRowNumber + 1, 2), pWs.Cells(codeRowNumber + 1, 2).End(xlDown))
            For Each headerCell In searchRange1
                For Each rowHeaderCell In searchRange2
                    Set dataCell = pWs.Cells(rowHeaderCell.row, headerCell.Column)
                    If IsEmpty(dataCell.value) Then
                        dataCell.value = 0
                    End If
                    DoEvents
                Next rowHeaderCell
                DoEvents
            Next headerCell
            
            
        End If
        DoEvents
    Next volCurve
    
End Sub