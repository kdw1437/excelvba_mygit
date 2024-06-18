VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "postCorrMissingData"
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
' @return {Worksheet} worksheet ��ü
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
' @return {Long} ���� Row ��
Public Property Get startRow() As Long
    startRow = pStartRow
End Property

' @property Let StartRow
' @param {Long} value - ���� Row �� Set
Public Property Let startRow(value As Long)
    pStartRow = value
End Property

' column �ĺ��� property
' @property Get Column
' @return {String} column �ĺ���
Public Property Get Column() As String
    Column = pColumn
End Property

' @property Let Column
' @param {String} value - column �ĺ��� Set
Public Property Let Column(value As String)
    pColumn = value
End Property

' worksheet data�κ��� JSON string�� �����.
'
' @method GenerateJSON
' @return {String} worksheet�� ���� ���� JSON string ������
Public Function GenerateJSON() As String
    Dim dataCollection As New Collection
    Dim i As Long
    i = pStartRow
    
    'Į�� �����ֱ�
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
        DoEvents
    Loop
    
    ' JsonConverter�� ����ؼ� collection�� JSON���� �ٲ��ش�.
    GenerateJSON = JsonConverter.ConvertToJson(dataCollection, Whitespace:=2)
End Function
