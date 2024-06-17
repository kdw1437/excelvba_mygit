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
' @usage - worksheet setter�̴�.
Public Property Set Ws(ByVal value As Worksheet)
    Set pWs = value
End Property

' Getter for pWs
' @method ws
' @return {Worksheet}
' @usage - worksheet getter�̴�.
Public Property Get Ws() As Worksheet
    Set Ws = pWs
End Property

' Getter for pArrayOfIndex
' @method ArrayOfIndex
' @return {Variant}
' @usage - array of index getter�̴�.
Public Property Get ArrayOfIndex() As Variant
    ArrayOfIndex = pArrayOfIndex
End Property

' Setter for pArrayOfIndex
' @method ArrayOfIndex
' @param {Variant} value
' @usage - array of index setter�̴�.
Public Property Let ArrayOfIndex(ByVal value As Variant)
    pArrayOfIndex = value
End Property

' Getter for pTitleCell
' @method TitleCell
' @return {Range}
' @usage - title cell�� getter�̴�.
Public Property Get TitleCell() As Range
    Set TitleCell = pTitleCell
End Property

' Setter for pTitleCell
' @method TitleCell
' @param {Range} value
' @usage - title cell�� setter�̴�.
Public Property Set TitleCell(ByVal value As Range)
    Set pTitleCell = value
End Property

' Setter for pStartCell
' @method startCell
' @param {Range} value
' @usage - start cell�� setter�̴�.
Public Property Set startCell(value As Range)
    Set pStartCell = value
End Property

' Getter for pStartCell
' @method startCell
' @return {Range}
' @usage - start cell�� getter�̴�.
Public Property Get startCell() As Range
    Set startCell = pStartCell
End Property

Private Sub Class_Initialize()
    ' pArrayOfIndex�� empty array�� �ʱ�ȭ�Ѵ�.
    ' � element�� array�� �߰��Ǳ� ������, array�� ���� �غ� ���ִ�.
    ReDim pArrayOfIndex(0 To 0)
End Sub
' @method PopulateArrayOfIndex
' @usage - pStartCell�� �ٰ��ؼ� pArrayOfIndex�� ä���.
Public Sub PopulateArrayOfIndex()
    Dim currentCell As Range
    Set currentCell = pStartCell ' class property pStartCell
    Dim cellCount As Integer
    cellCount = 0
    
    Do
        ' currency cell�� ������� Ȯ���Ѵ�.
        If IsEmpty(currentCell.value) Then
            Exit Do ' �� ���� �߰ߵǸ� loop�� ���� ������.
        End If
        
        ' array�� Resize�ϰ�, current cell�� ���� �ο��Ѵ�.
        cellCount = cellCount + 1
        ReDim Preserve pArrayOfIndex(1 To cellCount)
        pArrayOfIndex(cellCount) = currentCell.value
        
        ' ���������� 2 Į�� �̵��Ѵ�.
        Set currentCell = currentCell.Offset(0, 2)
    Loop
End Sub

' @method GenerateDataString2
' @return {String}
' @usage - array of index �����ͷκ��� JSON string�� �����Ѵ�.
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
            
            j = 1 ' inner loop counter�� �ʱ�ȭ�Ѵ�.
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
    
    ' jsonConverter�� ����ؼ� ��ü collection�� JSON string���� �����Ѵ�.
    GenerateDataString2 = JsonConverter.ConvertToJson(arrayOfIndexData)
End Function


