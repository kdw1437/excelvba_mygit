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
' @usage - Table1 range의 시작점을 set
Public Property Set Table1Point(value As Range)
    Set pTable1Point = value
End Property

' Getter for Table1Point
' @method Table1Point
' @return {Range}
' @usage - Table1 range의 시작점을 get
Public Property Get Table1Point() As Range
    Set Table1Point = pTable1Point
End Property

' Setter for fxRow
' @method fxRow
' @param {Range} value
' @usage - fxRow range를 set
Public Property Set fxRow(value As Range)
    Set pFxRow = value
End Property

' Getter for fxRow
' @method fxRow
' @return {Range}
' @usage - fxRow range를 get
Public Property Get fxRow() As Range
    Set fxRow = pFxRow
End Property

' Setter for Worksheet
' @method Worksheet
' @param {Worksheet} value
' @usage - worksheet의 setter
Public Property Set Worksheet(value As Worksheet)
    Set pWs = value
End Property

' Getter for Worksheet
' @method Worksheet
' @return {Worksheet}
' @usage - worksheet의 getter
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWs
End Property

' data의 JSON string을 만든다.
' @method GenerateJson2
' @return {String}
' @usage - Table1Point와 fxRow 사이에 있는 data로 JSON string을 만든다.
Public Function GenerateJson2() As String
    ' jsonArray를 만들기 위한 collection 생성. jsonObject는 jsonArray의 요소.
    Dim dataList As Collection
    Set dataList = New Collection
    
    Dim i As Integer
    For i = pTable1Point.row + 1 To pFxRow.row - 2
        ' jsonObject를 만들기 위한 dictionary 객체 생성.
        Dim dataDict As Object
        Set dataDict = CreateObject("Scripting.Dictionary")
        
        ' worksheet으로 부터 data회수
        dataDict("dataId") = pWs.Cells(i, pTable1Point.Column).value
        dataDict("price") = pWs.Cells(i, pTable1Point.Column + 1).value
        
        ' dictionary를 collection에 추가한다.
        dataList.Add dataDict
        DoEvents
    Next i
    
    ' collection을 JsonString으로 바꾸기 위해, ConvertToJson 메소드를 이용한다.
    GenerateJson2 = JsonConverter.ConvertToJson(dataList, Whitespace:=2)
End Function

