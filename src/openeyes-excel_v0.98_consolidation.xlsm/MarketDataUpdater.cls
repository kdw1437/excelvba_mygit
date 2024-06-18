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

' Worksheet property
' @method Worksheet
' @return {Worksheet}
' @usage - 설정된 Worksheet를 가져온다.
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWs
End Property

' @method Worksheet
' @param {Worksheet} wsTarget - 설정할 Worksheet
' @usage - Worksheet를 설정한다.
Public Property Set Worksheet(wsTarget As Worksheet)
    Set pWs = wsTarget
End Property

' PricesCollection property
' @method PricesCollection
' @return {Collection}
' @usage - 설정된 Prices 컬렉션을 가져온다.
Public Property Get PricesCollection() As Collection
    Set PricesCollection = pPrices
End Property

' @method PricesCollection
' @param {Collection} pricesCol - 설정할 Prices 컬렉션
' @usage - Prices 컬렉션을 설정한다.
Public Property Set PricesCollection(pricesCol As Collection)
    Set pPrices = pricesCol
End Property

' update를 수행하는 메소드
' @method UpdatePrices
' @usage - 'Equity' 행을 찾아 해당 데이터로 가격을 업데이트한다.
Public Sub UpdatePrices()
    Dim equityRow As Range
    Set equityRow = pWs.Columns(1).Find(What:="Equity", LookIn:=xlValues, Lookat:=xlPart)
    
    If Not equityRow Is Nothing Then
        Dim startRow As Integer
        startRow = equityRow.row + 4
        
        UpdatePriceForCodes startRow
    End If
End Sub

' 메소드를 캡슐화한다.
' @method UpdatePriceForCodes
' @param {Integer} StartRow - 시작 행 번호
' @usage - 코드별로 가격을 업데이트한다.
Private Sub UpdatePriceForCodes(startRow As Integer)
    Dim codeCell As Range
    Dim codeValue As String
    
    For Each codeCell In pWs.Range(pWs.Cells(startRow, 1), pWs.Cells(pWs.Rows.Count, 1))
        
        If IsEmpty(codeCell.value) Then Exit For

        codeValue = codeCell.value

        ' 맞는 데이터가 있을 시, 가격을 업데이트 시켜준다.
        UpdatePriceIfMatchFound codeCell, codeValue
    Next codeCell
End Sub

' 지표가 codeValue와 일치할 시, 가격을 지표 옆에 넣어주는 코드
' @method UpdatePriceIfMatchFound
' @param {Range} codeCell - 코드 셀
' @param {String} codeValue - 코드 값
' @usage - 일치하는 데이터가 있을 때 가격을 업데이트한다.
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

' 모든 단계를 결합하여 업데이트를 수행하는 메소드
' @method UpdatePrices2
' @usage - 'Equity' 행을 찾아 해당 데이터로 가격을 업데이트한다.
Public Sub UpdatePrices2()
    Dim equityRow As Range
    Set equityRow = pWs.Columns(1).Find(What:="Equity", LookIn:=xlValues, Lookat:=xlPart)
    
    If Not equityRow Is Nothing Then
        Dim startRow As Integer
        startRow = equityRow.row + 4
        
        Dim codeCell As Range
        Dim codeValue As String
        Dim priceInfo As Variant
        Dim i As Integer
        
        For Each codeCell In pWs.Range(pWs.Cells(startRow, 1), pWs.Cells(pWs.Rows.Count, 1))
            If IsEmpty(codeCell.value) Then Exit For

            codeValue = codeCell.value

            ' 맞는 dataId를 check하고 가격을 update한다.
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

' 최적화된 방법으로 가격을 업데이트하는 메소드
' @method UpdatePricesOptimized
' @usage - Dictionary 객체를 사용하여 빠르게 가격을 업데이트한다.
Public Sub UpdatePricesOptimized()
    Dim equityRow As Range
    Set equityRow = pWs.Columns(1).Find(What:="Equity", LookIn:=xlValues, Lookat:=xlPart)
    
    If Not equityRow Is Nothing Then
        Dim startRow As Integer
        startRow = equityRow.row + 4
        
        Dim lastRow As Long
        lastRow = pWs.Cells(startRow, 1).End(xlDown).row
        
        Dim codeCell As Range
        Dim codeValue As String
        Dim priceDict As Object
        Set priceDict = CreateObject("Scripting.Dictionary")
        
        Dim priceInfo As Variant
        Dim i As Integer
        
        ' 빠른 lookup을 위해 pPrices collection을 Dictionary로 변경한다.
        For i = 1 To pPrices.Count
            Set priceInfo = pPrices(i)
            priceDict(priceInfo("dataId")) = priceInfo("price")
        Next i
        
        'range를 가지기 위해 변수 선언
        Dim codeRange As Range
        ' range로 부터 읽어낸 데이터를 저장하기 위해 변수 선언
        Dim dataArray As Variant
        ' codeRange를 특정 range로 선언
        Set codeRange = pWs.Range(pWs.Cells(startRow, 1), pWs.Cells(lastRow, 1))
        ' 특정 range로부터 값을 읽어와서 dataArray에 저장
        dataArray = codeRange.value
        
        ' 배열을 통해 반복하고 가격을 업데이트하기
        For i = 1 To UBound(dataArray, 1)
            codeValue = dataArray(i, 1)
            If IsEmpty(codeValue) Then Exit For
            
            If priceDict.Exists(codeValue) Then
                pWs.Cells(startRow + i - 1, 2).value = priceDict(codeValue)
            End If
            
            DoEvents
        Next i
    End If
End Sub
