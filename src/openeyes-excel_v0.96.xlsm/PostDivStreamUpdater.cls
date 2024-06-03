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
Public Property Set ws(ByVal value As Worksheet)
    Set pWs = value
End Property

' Getter for pWs
Public Property Get ws() As Worksheet
    Set ws = pWs
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
    ' pArrayOfIndex�� empty array�� �ʱ�ȭ�Ѵ�.
    ' � element�� array�� �߰��Ǳ� ������, array�� ���� �غ� ���ִ�.
    ReDim pArrayOfIndex(0 To 0)
End Sub
' pStartCell�� �ٰ��ؼ� pArrayOfIndex�� ä���.
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

' DataString�� ����� �޼ҵ�
Public Function GenerateDataString() As String
    Dim arraySize As Integer
    arraySize = UBound(pArrayOfIndex) ' ArrayofIndex�� ������
    Dim DataString As String
    DataString = "["
    
    If arraySize > 0 Then
        Dim i As Integer, j As Integer
        Dim IndexName As String, DateTime As String, value As Double
        Dim dividendsArray As String
        
        For i = 1 To arraySize
            IndexName = pArrayOfIndex(i) ' ArrayofIndex�� ����(String)�� IndexName ���������� �޴´�.
            j = 1 ' inner loop counter�� �ʱ�ȭ�Ѵ�.
            dividendsArray = "["
            
            Do While Not IsEmpty(pTitleCell.Offset(3 + j, (i - 1) * 2).value) ' property TitleCell�� ����Ѵ�.
                DateTime = pTitleCell.Offset(3 + j, (i - 1) * 2).value
                value = pTitleCell.Offset(3 + j, (i - 1) * 2 + 1).value
                
                
                ' yield object�� �����, �� ���� yieldsArray�� �߰��Ѵ�.
                If dividendsArray <> "[" Then dividendsArray = dividendsArray & ","
                dividendsArray = dividendsArray & "{""date"": """ & DateTime & """, ""value"": " & value & "}"
                
                j = j + 1
            Loop
            dividendsArray = dividendsArray & "]"
            
            ' main JSON object�� Interest�̸��� Yields array�� �߰��Ѵ�.
            If i > 1 Then DataString = DataString & ","
            DataString = DataString & "{" & _
                                         """dataId"": """ & IndexName & """, " & _
                                         """dividends"": " & dividendsArray & "}"
        Next i
    End If
    DataString = DataString & "]"
    
    GenerateDataString = DataString ' DataString�� �Լ��� output���� return�Ѵ�.
End Function


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
            
            j = 1 ' Initialize inner loop counter
            Do While Not IsEmpty(pTitleCell.Offset(3 + j, (i - 1) * 2).value)
                Dim dividend As Dictionary
                Set dividend = New Dictionary
                dividend.Add "date", CStr(pTitleCell.Offset(3 + j, (i - 1) * 2).value)
                dividend.Add "value", pTitleCell.Offset(3 + j, (i - 1) * 2 + 1).value
                
                dividends.Add dividend
                j = j + 1
            Loop
            
            indexData.Add "dividends", dividends
            arrayOfIndexData.Add indexData
        Next i
    End If
    
    ' Convert the entire collection to a JSON string using jsonConverter
    GenerateDataString2 = JsonConverter.ConvertToJson(arrayOfIndexData)
End Function


