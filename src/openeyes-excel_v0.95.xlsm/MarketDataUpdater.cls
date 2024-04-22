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

' update�� �����ϴ� �޼ҵ�
Public Sub UpdatePrices()
    Dim equityRow As Range
    Set equityRow = pWs.Columns(1).Find(What:="Equity", LookIn:=xlValues, LookAt:=xlPart)
    
    If Not equityRow Is Nothing Then
        Dim StartRow As Integer
        StartRow = equityRow.row + 4
        
        UpdatePriceForCodes StartRow
    End If
End Sub

' �޼ҵ带 ĸ��ȭ�Ѵ�.
Private Sub UpdatePriceForCodes(StartRow As Integer)
    Dim codeCell As Range
    Dim codeValue As String
    
    For Each codeCell In pWs.Range(pWs.Cells(StartRow, 1), pWs.Cells(pWs.Rows.Count, 1))
        
        If IsEmpty(codeCell.value) Then Exit For

        codeValue = codeCell.value

        ' �´� �����Ͱ� ���� ��, ������ ������Ʈ �����ش�.
        UpdatePriceIfMatchFound codeCell, codeValue
    Next codeCell
End Sub

' ��ǥ�� codeValue�� ��ġ�� ��, ������ ��ǥ ���� �־��ִ� �ڵ�
Private Sub UpdatePriceIfMatchFound(codeCell As Range, codeValue As String)
    Dim priceInfo As Variant
    Dim i As Integer
    
    For i = 1 To pPrices.Count
        Set priceInfo = pPrices(i)

        
        If priceInfo("dataId") = codeValue Then
            
            codeCell.Offset(0, 1).value = priceInfo("price")
            Exit For
        End If
    Next i
End Sub
