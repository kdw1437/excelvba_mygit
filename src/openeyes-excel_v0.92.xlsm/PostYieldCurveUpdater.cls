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
Public Property Get ArrayOfCurrency() As Variant
    ArrayOfCurrency = pArrayOfCurrency
End Property

' Setter for pArrayOfCurrency
Public Property Let ArrayOfCurrency(ByVal Value As Variant)
    pArrayOfCurrency = Value
End Property



' Getter for pYieldCurveRow
Public Property Get YieldCurveRow() As Range
    Set YieldCurveRow = pYieldCurveRow
End Property

' Setter for pYieldCurveRow
Public Property Set YieldCurveRow(ByVal Value As Range)
    Set pYieldCurveRow = Value
End Property


' Setter for pStartCell
Public Property Set startCell(Value As Range)
    Set pStartCell = Value
End Property

' Getter for pStartCell
Public Property Get startCell() As Range
    Set startCell = pStartCell
End Property
Private Sub Class_Initialize()
    ' Initialize the pArrayOfCurrency as an empty array
    ' This ensures that the array is ready to be used, even before any elements are added to it
    ReDim pArrayOfCurrency(0 To 0)
End Sub
' pStartCell에 근거해서 pArrayOfCurrency를 채운다.
Public Sub PopulateArrayOfCurrency()
    Dim currentCell As Range
    Set currentCell = pStartCell ' class property pStartCell
    Dim cellCount As Integer
    cellCount = 0
    
    Do
        ' currency cell이 비었는지 확인한다.
        If IsEmpty(currentCell.Value) Then
            Exit Do ' 빈 셀이 발견되면 loop를 빠져 나간다.
        End If
        
        ' array를 Resize하고, current cell의 값을 부여한다.
        cellCount = cellCount + 1
        ReDim Preserve pArrayOfCurrency(1 To cellCount)
        pArrayOfCurrency(cellCount) = currentCell.Value
        
        ' 오른쪽으로 2 칼럼 이동한다.
        Set currentCell = currentCell.Offset(0, 2)
    Loop
End Sub


' DataString을 만드는 메소드
Public Function GenerateDataString() As String
    Dim arraySize As Integer
    arraySize = UBound(pArrayOfCurrency) ' ArrayofCurrency의 사이즈
    Dim DataString As String
    DataString = "["
    
    If arraySize > 0 Then
        Dim i As Integer, j As Integer
        Dim InterestName As String, tenor As Double, Rate As Double, RiskCode As String
        Dim yieldsArray As String
        
        For i = 1 To arraySize
            InterestName = pArrayOfCurrency(i) ' ArrayofCurrency의 원소(String)을 InterestName 참조변수로 받는다.
            j = 1 ' inner loop counter를 초기화한다.
            yieldsArray = "["
            
            Do While Not IsEmpty(pYieldCurveRow.Offset(3 + j, (i - 1) * 2).Value) ' property YieldCurveRow를 사용한다.
                tenor = pYieldCurveRow.Offset(3 + j, (i - 1) * 2).Value
                Rate = pYieldCurveRow.Offset(3 + j, (i - 1) * 2 + 1).Value
                RiskCode = Format(tenor * 360, "00000")
                
                ' yield object를 만들고, 이 것을 yieldsArray에 추가한다.
                If yieldsArray <> "[" Then yieldsArray = yieldsArray & ","
                yieldsArray = yieldsArray & "{""tenor"": " & tenor & ", ""rate"": " & Rate & "}"
                
                j = j + 1
            Loop
            yieldsArray = yieldsArray & "]"
            
            ' main JSON object에 Interest이름과 Yields array를 추가한다.
            If i > 1 Then DataString = DataString & ","
            DataString = DataString & "{" & _
                                         """dataId"": """ & InterestName & """, " & _
                                         """currency"": """ & Left(InterestName, 3) & """, " & _
                                         """yields"": " & yieldsArray & "}"
        Next i
    End If
    DataString = DataString & "]"
    
    GenerateDataString = DataString ' DataString을 함수의 output으로 return한다.
End Function



