VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostPriceUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pTable1Point As Range
Private pFxRow As Range
Private pWs As Worksheet

' Setter for Table1Point
' @method Table1Point
' @param {Range} value
' @usage - Table1 range�� �������� set
Public Property Set Table1Point(value As Range)
    Set pTable1Point = value
End Property

' Getter for Table1Point
' @method Table1Point
' @return {Range}
' @usage - Table1 range�� �������� get
Public Property Get Table1Point() As Range
    Set Table1Point = pTable1Point
End Property

' Setter for fxRow
' @method fxRow
' @param {Range} value
' @usage - fxRow range�� set
Public Property Set fxRow(value As Range)
    Set pFxRow = value
End Property

' Getter for fxRow
' @method fxRow
' @return {Range}
' @usage - fxRow range�� get
Public Property Get fxRow() As Range
    Set fxRow = pFxRow
End Property

' Setter for Worksheet
' @method Worksheet
' @param {Worksheet} value
' @usage - worksheet�� setter
Public Property Set Worksheet(value As Worksheet)
    Set pWs = value
End Property

' Getter for Worksheet
' @method Worksheet
' @return {Worksheet}
' @usage - worksheet�� getter
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWs
End Property

' data�� JSON string�� �����.
' @method GenerateJson2
' @return {String}
' @usage - Table1Point�� fxRow ���̿� �ִ� data�� JSON string�� �����.
Public Function GenerateJson2() As String
    ' jsonArray�� ����� ���� collection ����. jsonObject�� jsonArray�� ���.
    Dim dataList As Collection
    Set dataList = New Collection
    
    Dim i As Integer
    For i = pTable1Point.row + 1 To pFxRow.row - 2
        ' jsonObject�� ����� ���� dictionary ��ü ����.
        Dim dataDict As Object
        Set dataDict = CreateObject("Scripting.Dictionary")
        
        ' worksheet���� ���� dataȸ��
        dataDict("dataId") = pWs.Cells(i, pTable1Point.Column).value
        dataDict("price") = pWs.Cells(i, pTable1Point.Column + 1).value
        
        ' dictionary�� collection�� �߰��Ѵ�.
        dataList.Add dataDict
        DoEvents
    Next i
    
    ' collection�� JsonString���� �ٲٱ� ����, ConvertToJson �޼ҵ带 �̿��Ѵ�.
    GenerateJson2 = JsonConverter.ConvertToJson(dataList, Whitespace:=2)
End Function
