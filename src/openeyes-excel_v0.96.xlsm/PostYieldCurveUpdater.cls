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
Public Property Let ArrayOfCurrency(ByVal value As Variant)
    pArrayOfCurrency = value
End Property



' Getter for pYieldCurveRow
Public Property Get YieldCurveRow() As Range
    Set YieldCurveRow = pYieldCurveRow
End Property

' Setter for pYieldCurveRow
Public Property Set YieldCurveRow(ByVal value As Range)
    Set pYieldCurveRow = value
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
    ' Initialize the pArrayOfCurrency as an empty array
    ' This ensures that the array is ready to be used, even before any elements are added to it
    ReDim pArrayOfCurrency(0 To 0)
End Sub
' pStartCell�� �ٰ��ؼ� pArrayOfCurrency�� ä���.
Public Sub PopulateArrayOfCurrency()
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
        ReDim Preserve pArrayOfCurrency(1 To cellCount)
        pArrayOfCurrency(cellCount) = currentCell.value
        
        ' ���������� 2 Į�� �̵��Ѵ�.
        Set currentCell = currentCell.Offset(0, 2)
    Loop
End Sub


' DataString�� ����� �޼ҵ�
Public Function GenerateDataString() As String
    Dim arraySize As Integer
    arraySize = UBound(pArrayOfCurrency) ' ArrayofCurrency�� ������
    Dim DataString As String
    DataString = "["
    
    If arraySize > 0 Then
        Dim i As Integer, j As Integer
        Dim InterestName As String, tenor As Double, Rate As Double, RiskCode As String
        Dim yieldsArray As String
        
        For i = 1 To arraySize
            InterestName = pArrayOfCurrency(i) ' ArrayofCurrency�� ����(String)�� InterestName ���������� �޴´�.
            j = 1 ' inner loop counter�� �ʱ�ȭ�Ѵ�.
            yieldsArray = "["
            
            Do While Not IsEmpty(pYieldCurveRow.Offset(3 + j, (i - 1) * 2).value) ' property YieldCurveRow�� ����Ѵ�.
                tenor = pYieldCurveRow.Offset(3 + j, (i - 1) * 2).value
                Rate = pYieldCurveRow.Offset(3 + j, (i - 1) * 2 + 1).value
                RiskCode = Format(tenor * 360, "00000")
                
                ' yield object�� �����, �� ���� yieldsArray�� �߰��Ѵ�.
                If yieldsArray <> "[" Then yieldsArray = yieldsArray & ","
                yieldsArray = yieldsArray & "{""tenor"": " & tenor & ", ""rate"": " & Rate & "}"
                
                j = j + 1
            Loop
            yieldsArray = yieldsArray & "]"
            
            ' main JSON object�� Interest�̸��� Yields array�� �߰��Ѵ�.
            If i > 1 Then DataString = DataString & ","
            DataString = DataString & "{" & _
                                         """dataId"": """ & InterestName & """, " & _
                                         """currency"": """ & Left(InterestName, 3) & """, " & _
                                         """yields"": " & yieldsArray & "}"
        Next i
    End If
    DataString = DataString & "]"
    
    GenerateDataString = DataString ' DataString�� �Լ��� output���� return�Ѵ�.
End Function


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
            Loop

            Dim currencyData As New Dictionary
            Set currencyData = New Dictionary
            
            currencyData.Add "dataId", InterestName
            currencyData.Add "currency", Left(InterestName, 3)
            currencyData.Add "yields", yieldsArray
            dataArray.Add currencyData
        Next i
         
    End If
    
    ' Convert the Collection of Dictionaries to JSON
    GenerateDataString2 = JsonConverter.ConvertToJson(dataArray)
End Function
