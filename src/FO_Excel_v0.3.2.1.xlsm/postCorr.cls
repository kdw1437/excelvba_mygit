VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "postCorr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class postCorr
Private pWorksheet As Worksheet
Private pStartRow As Long
Private pColumn As String

' workSheet property
' @property Get Worksheet
' @return {Worksheet} worksheet 객체
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWorksheet
End Property

' @property Set Worksheet
' @param {Worksheet} value
Public Property Set Worksheet(value As Worksheet)
    Set pWorksheet = value
End Property

' starting row Property
' @property Get StartRow
' @return {Long} 시작 Row 값
Public Property Get StartRow() As Long
    StartRow = pStartRow
End Property

' @property Let StartRow
' @param {Long} value - 시작 Row 값 Set
Public Property Let StartRow(value As Long)
    pStartRow = value
End Property

' column 식별자 property
' @property Get Column
' @return {String} column 식별자
Public Property Get Column() As String
    Column = pColumn
End Property

' @property Let Column
' @param {String} value - column 식별자 Set
Public Property Let Column(value As String)
    pColumn = value
End Property

' worksheet data로부터 JSON string을 만든다.
'
' @method GenerateJSON
' @return {String} worksheet로 부터 만든 JSON string 데이터
Public Function GenerateJSON() As String
    Dim dataCollection As New Collection
    Dim i As Long
    i = pStartRow
    
    '칼럼 정해주기
    Dim colE As String
    Dim colF As String
    Dim colG As String

    colE = pColumn
    colF = Chr(Asc(pColumn) + 1)
    colG = Chr(Asc(pColumn) + 2)
    
    Do While pWorksheet.Cells(i, colE).value <> ""
        Dim dataDict As Object
        Set dataDict = CreateObject("Scripting.Dictionary")
        
        dataDict.Add "dataId", pWorksheet.Cells(i, colE).value & ":" & pWorksheet.Cells(i, colF).value
        dataDict.Add "dataId1", pWorksheet.Cells(i, colE).value
        dataDict.Add "dataId2", pWorksheet.Cells(i, colF).value
        dataDict.Add "corr", pWorksheet.Cells(i, colG).value
        
        dataCollection.Add dataDict
        i = i + 1
    Loop
    
    ' JsonConverter를 사용해서 collection을 JSON으로 바꿔준다.
    GenerateJSON = JsonConverter.ConvertToJson(dataCollection, Whitespace:=2)
End Function

