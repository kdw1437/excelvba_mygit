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

' Combined method to perform the update
Public Sub UpdatePrices2()
    Dim equityRow As Range
    Set equityRow = pWs.Columns(1).Find(What:="Equity", LookIn:=xlValues, LookAt:=xlPart)
    
    If Not equityRow Is Nothing Then
        Dim StartRow As Integer
        StartRow = equityRow.row + 4
        
        Dim codeCell As Range
        Dim codeValue As String
        Dim priceInfo As Variant
        Dim i As Integer
        
        For Each codeCell In pWs.Range(pWs.Cells(StartRow, 1), pWs.Cells(pWs.Rows.Count, 1))
            If IsEmpty(codeCell.value) Then Exit For

            codeValue = codeCell.value

            ' Check for matching dataId and update price
            For i = 1 To pPrices.Count
                Set priceInfo = pPrices(i)

                If priceInfo("dataId") = codeValue Then
                    codeCell.Offset(0, 1).value = priceInfo("price")
                    Exit For
                End If
            Next i
        Next codeCell
    End If
End Sub

Public Sub UpdatePricesOptimized()
    Dim equityRow As Range
    Set equityRow = pWs.Columns(1).Find(What:="Equity", LookIn:=xlValues, LookAt:=xlPart)
    
    If Not equityRow Is Nothing Then
        Dim StartRow As Integer
        StartRow = equityRow.row + 4
        
        Dim lastRow As Long
        lastRow = pWs.Cells(StartRow, 1).End(xlDown).row
        
        Dim codeCell As Range
        Dim codeValue As String
        Dim priceDict As Object
        Set priceDict = CreateObject("Scripting.Dictionary")
        
        Dim priceInfo As Variant
        Dim i As Integer
        
        ' Convert pPrices collection to Dictionary for faster lookup
        For i = 1 To pPrices.Count
            Set priceInfo = pPrices(i)
            priceDict(priceInfo("dataId")) = priceInfo("price")
        Next i
        
        ' Read the range into a VBA array
        Dim codeRange As Range
        Dim dataArray As Variant
        Set codeRange = pWs.Range(pWs.Cells(StartRow, 1), pWs.Cells(lastRow, 1))
        dataArray = codeRange.value
        
        ' Iterate through the array and update prices
        For i = 1 To UBound(dataArray, 1)
            codeValue = dataArray(i, 1)
            If IsEmpty(codeValue) Then Exit For
            
            If priceDict.Exists(codeValue) Then
                pWs.Cells(StartRow + i - 1, 2).value = priceDict(codeValue)
            End If
        Next i
    End If
End Sub