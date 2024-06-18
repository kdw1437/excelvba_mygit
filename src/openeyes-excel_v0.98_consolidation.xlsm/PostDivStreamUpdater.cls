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
' @method ws
' @param {Worksheet} value
' @usage - worksheet setter이다.
Public Property Set ws(ByVal value As Worksheet)
    Set pWs = value
End Property

' Getter for pWs
' @method ws
' @return {Worksheet}
' @usage - worksheet getter이다.
Public Property Get ws() As Worksheet
    Set ws = pWs
End Property

' Getter for pArrayOfIndex
' @method ArrayOfIndex
' @return {Variant}
' @usage - array of index getter이다.
Public Property Get ArrayOfIndex() As Variant
    ArrayOfIndex = pArrayOfIndex
End Property

' Setter for pArrayOfIndex
' @method ArrayOfIndex
' @param {Variant} value
' @usage - array of index setter이다.
Public Property Let ArrayOfIndex(ByVal value As Variant)
    pArrayOfIndex = value
End Property

' Getter for pTitleCell
' @method TitleCell
' @return {Range}
' @usage - title cell의 getter이다.
Public Property Get TitleCell() As Range
    Set TitleCell = pTitleCell
End Property

' Setter for pTitleCell
' @method TitleCell
' @param {Range} value
' @usage - title cell의 setter이다.
Public Property Set TitleCell(ByVal value As Range)
    Set pTitleCell = value
End Property

' Setter for pStartCell
' @method startCell
' @param {Range} value
' @usage - start cell의 setter이다.
Public Property Set StartCell(value As Range)
    Set pStartCell = value
End Property

' Getter for pStartCell
' @method startCell
' @return {Range}
' @usage - start cell의 getter이다.
Public Property Get StartCell() As Range
    Set StartCell = pStartCell
End Property

Private Sub Class_Initialize()
    ' pArrayOfIndex를 empty array로 초기화한다.
    ' 어떤 element가 array에 추가되기 전에도, array가 사용될 준비가 되있다.
    ReDim pArrayOfIndex(0 To 0)
End Sub
' @method PopulateArrayOfIndex
' @usage - pStartCell에 근거해서 pArrayOfIndex를 채운다.
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

' @method GenerateDataString2
' @return {String}
' @usage - array of index 데이터로부터 JSON string을 생성한다.
Public Function GenerateDataString2() As String
    Dim arrayOfIndexData As Collection
    Set arrayOfIndexData = New Collection
    
    Dim arraySize As Integer
    arraySize = UBound(pArrayOfIndex)
    
    If arraySize > 0 Then
        Dim i As Integer, j As Integer
        Dim indexData As Dictionary
        
        For i = 1 To arraySize
            Set indexData = New Dictionary
            indexData.Add "dataId", pArrayOfIndex(i)
            
            Dim dividends As Collection
            Set dividends = New Collection
            
            j = 1 ' inner loop counter를 초기화한다.
            Do While Not IsEmpty(pTitleCell.Offset(3 + j, (i - 1) * 2).value)
                Dim dividend As Dictionary
                Set dividend = New Dictionary
                dividend.Add "date", CStr(pTitleCell.Offset(3 + j, (i - 1) * 2).value)
                dividend.Add "value", pTitleCell.Offset(3 + j, (i - 1) * 2 + 1).value
                
                dividends.Add dividend
                j = j + 1
                DoEvents
            Loop
            
            indexData.Add "dividends", dividends
            arrayOfIndexData.Add indexData
            DoEvents
        Next i
    End If
    
    ' jsonConverter를 사용해서 전체 collection을 JSON string으로 변경한다.
    GenerateDataString2 = JsonConverter.ConvertToJson(arrayOfIndexData)
End Function



