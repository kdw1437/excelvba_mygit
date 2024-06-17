VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostYieldCurveUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pArrayOfCurrency As Variant
Private pYieldCurveRow As Range
Private pStartCell As Range

' Getter for pArrayOfCurrency
' @method ArrayOfCurrency
' @return {Variant}
' @usage - pArrayOfCurrency 값을 가져온다.
Public Property Get ArrayOfCurrency() As Variant
    ArrayOfCurrency = pArrayOfCurrency
End Property

' Setter for pArrayOfCurrency
' @method ArrayOfCurrency
' @param {Variant} value
' @usage - pArrayOfCurrency 값을 설정한다.
Public Property Let ArrayOfCurrency(ByVal value As Variant)
    pArrayOfCurrency = value
End Property

' Getter for pYieldCurveRow
' @method YieldCurveRow
' @return {Range}
' @usage - pYieldCurveRow 값을 가져온다.
Public Property Get YieldCurveRow() As Range
    Set YieldCurveRow = pYieldCurveRow
End Property

' Setter for pYieldCurveRow
' @method YieldCurveRow
' @param {Range} value
' @usage - pYieldCurveRow 값을 설정한다.
Public Property Set YieldCurveRow(ByVal value As Range)
    Set pYieldCurveRow = value
End Property

' Setter for pStartCell
' @method startCell
' @param {Range} value
' @usage - pStartCell 값을 설정한다.
Public Property Set startCell(value As Range)
    Set pStartCell = value
End Property

' Getter for pStartCell
' @method startCell
' @return {Range}
' @usage - pStartCell 값을 가져온다.
Public Property Get startCell() As Range
    Set startCell = pStartCell
End Property

' 클래스 초기화 시 실행되는 메서드
' @method Class_Initialize
' @usage - pArrayOfCurrency를 빈 배열로 초기화한다.
Private Sub Class_Initialize()
    ' Initialize the pArrayOfCurrency as an empty array
    ' This ensures that the array is ready to be used, even before any elements are added to it
    ReDim pArrayOfCurrency(0 To 0)
End Sub

' pStartCell에 근거해서 pArrayOfCurrency를 채운다.
' @method PopulateArrayOfCurrency
' @usage - pStartCell을 시작으로 pArrayOfCurrency를 채운다.
Public Sub PopulateArrayOfCurrency()
    Dim currentCell As Range
    Set currentCell = pStartCell ' class property pStartCell
    Dim cellCount As Integer
    cellCount = 0
    
    Do
        ' currency cell이 비었는지 확인한다.
        If IsEmpty(currentCell.value) Then
            Exit Do ' 빈 셀이 발견되면 loop를 빠져 나간다.
        End If
        
        ' array를 Resize하고, current cell의 값을 부여한다.
        cellCount = cellCount + 1
        ReDim Preserve pArrayOfCurrency(1 To cellCount)
        pArrayOfCurrency(cellCount) = currentCell.value
        
        ' 오른쪽으로 2 칼럼 이동한다.
        Set currentCell = currentCell.Offset(0, 2)
    Loop
End Sub

' 데이터를 JSON 형식의 문자열로 생성한다.
' @method GenerateDataString2
' @return {String}
' @usage - pArrayOfCurrency와 pYieldCurveRow 데이터를 JSON 문자열로 변환한다.
Public Function GenerateDataString2() As String
    Dim dataArray As New Collection
    Dim arraySize As Integer
    arraySize = UBound(pArrayOfCurrency)
    
    If arraySize > 0 Then
        Dim i As Integer, j As Integer
        Dim InterestName As String
        Dim yieldsArray As Collection
        Dim yieldObject As Dictionary
        
        For i = 1 To arraySize
            InterestName = pArrayOfCurrency(i)
            Set yieldsArray = New Collection

            j = 1
            Do While Not IsEmpty(pYieldCurveRow.Offset(3 + j, (i - 1) * 2).value)
                Set yieldObject = New Dictionary
                yieldObject.Add "tenor", pYieldCurveRow.Offset(3 + j, (i - 1) * 2).value
                yieldObject.Add "rate", pYieldCurveRow.Offset(3 + j, (i - 1) * 2 + 1).value
                yieldsArray.Add yieldObject
                j = j + 1
                DoEvents
            Loop

            Dim currencyData As New Dictionary
            Set currencyData = New Dictionary
            
            currencyData.Add "dataId", InterestName
            currencyData.Add "currency", Left(InterestName, 3)
            currencyData.Add "yields", yieldsArray
            dataArray.Add currencyData
            DoEvents
        Next i
         
    End If
    
    ' Dictionary의 collection을 JSON으로 변경한다.
    GenerateDataString2 = JsonConverter.ConvertToJson(dataArray)
End Function

