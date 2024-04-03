VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostDivStreamUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pArrayOfIndex As Variant
Private pTitleCell As Range
Private pStartCell As Range
Private pWs As Worksheet

' Setter for pWs
Public Property Set Ws(ByVal value As Worksheet)
    Set pWs = value
End Property

' Getter for pWs
Public Property Get Ws() As Worksheet
    Set Ws = pWs
End Property

' Getter for pArrayOfIndex
Public Property Get ArrayOfIndex() As Variant
    ArrayOfIndex = pArrayOfIndex
End Property

' Setter for pArrayOfIndex
Public Property Let ArrayOfIndex(ByVal value As Variant)
    pArrayOfIndex = value
End Property

' Getter for pTitleCell
Public Property Get TitleCell() As Range
    Set TitleCell = pTitleCell
End Property

' Setter for pTitleCell
Public Property Set TitleCell(ByVal value As Range)
    Set pTitleCell = value
End Property

' Setter for pStartCell
Public Property Set startCell(value As Range)
    Set pStartCell = value
End Property

' Getter for pStartCell
Public Property Get startCell() As Range
    Set startCell = pStartCell
End Property

Private Sub Class_Initialize()
    ' pArrayOfIndex를 empty array로 초기화한다.
    ' 어떤 element가 array에 추가되기 전에도, array가 사용될 준비가 되있다.
    ReDim pArrayOfIndex(0 To 0)
End Sub
' pStartCell에 근거해서 pArrayOfIndex를 채운다.
Public Sub PopulateArrayOfIndex()
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
        ReDim Preserve pArrayOfIndex(1 To cellCount)
        pArrayOfIndex(cellCount) = currentCell.value
        
        ' 오른쪽으로 2 칼럼 이동한다.
        Set currentCell = currentCell.Offset(0, 2)
    Loop
End Sub

' DataString을 만드는 메소드
Public Function GenerateDataString() As String
    Dim arraySize As Integer
    arraySize = UBound(pArrayOfIndex) ' ArrayofIndex의 사이즈
    Dim DataString As String
    DataString = "["
    
    If arraySize > 0 Then
        Dim i As Integer, j As Integer
        Dim IndexName As String, DateTime As String, value As Double
        Dim dividendsArray As String
        
        For i = 1 To arraySize
            IndexName = pArrayOfIndex(i) ' ArrayofIndex의 원소(String)을 IndexName 참조변수로 받는다.
            j = 1 ' inner loop counter를 초기화한다.
            dividendsArray = "["
            
            Do While Not IsEmpty(pTitleCell.Offset(3 + j, (i - 1) * 2).value) ' property TitleCell를 사용한다.
                DateTime = pTitleCell.Offset(3 + j, (i - 1) * 2).value
                value = pTitleCell.Offset(3 + j, (i - 1) * 2 + 1).value
                
                
                ' yield object를 만들고, 이 것을 yieldsArray에 추가한다.
                If dividendsArray <> "[" Then dividendsArray = dividendsArray & ","
                dividendsArray = dividendsArray & "{""date"": """ & DateTime & """, ""value"": " & value & "}"
                
                j = j + 1
            Loop
            dividendsArray = dividendsArray & "]"
            
            ' main JSON object에 Interest이름과 Yields array를 추가한다.
            If i > 1 Then DataString = DataString & ","
            DataString = DataString & "{" & _
                                         """dataId"": """ & IndexName & """, " & _
                                         """dividends"": " & dividendsArray & "}"
        Next i
    End If
    DataString = DataString & "]"
    
    GenerateDataString = DataString ' DataString을 함수의 output으로 return한다.
End Function





