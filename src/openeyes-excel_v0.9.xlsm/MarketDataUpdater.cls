VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MarketDataUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pWs As Worksheet
Private pPrices As Collection

Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWs
End Property

Public Property Set Worksheet(wsTarget As Worksheet)
    Set pWs = wsTarget
End Property

Public Property Get PricesCollection() As Collection
    Set PricesCollection = pPrices
End Property

Public Property Set PricesCollection(pricesCol As Collection)
    Set pPrices = pricesCol
End Property

' update를 수행하는 메소드
Public Sub UpdatePrices()
    Dim equityRow As Range
    Set equityRow = pWs.Columns(1).Find(What:="Equity", LookIn:=xlValues, LookAt:=xlPart)
    
    If Not equityRow Is Nothing Then
        Dim StartRow As Integer
        StartRow = equityRow.Row + 4
        
        UpdatePriceForCodes StartRow
    End If
End Sub

' 메소드를 캡슐화한다.
Private Sub UpdatePriceForCodes(StartRow As Integer)
    Dim codeCell As Range
    Dim codeValue As String
    
    For Each codeCell In pWs.Range(pWs.Cells(StartRow, 1), pWs.Cells(pWs.Rows.Count, 1))
        
        If IsEmpty(codeCell.Value) Then Exit For

        codeValue = codeCell.Value

        ' 맞는 데이터가 있을 시, 가격을 업데이트 시켜준다.
        UpdatePriceIfMatchFound codeCell, codeValue
    Next codeCell
End Sub

' 지표가 codeValue와 일치할 시, 가격을 지표 옆에 넣어주는 코드
Private Sub UpdatePriceIfMatchFound(codeCell As Range, codeValue As String)
    Dim priceInfo As Variant
    Dim i As Integer
    
    For i = 1 To pPrices.Count
        Set priceInfo = pPrices(i)

        
        If priceInfo("dataId") = codeValue Then
            
            codeCell.Offset(0, 1).Value = priceInfo("price")
            Exit For
        End If
    Next i
End Sub

