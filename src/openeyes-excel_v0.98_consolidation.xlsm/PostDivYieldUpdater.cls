VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostDivYieldUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pWs As Worksheet
Private pDivCell As Range
Private pdataIdRange As Range

' Getter for Worksheet
' @method Worksheet
' @return {Worksheet}
' @usage - worksheet�� getter
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWs
End Property

' Setter for Worksheet
' @method Worksheet
' @param {Worksheet} value
' @usage - worksheet�� setter
Public Property Set Worksheet(value As Worksheet)
    Set pWs = value
End Property

' Getter for DivCell Range
' @method DivCell
' @return {Range}
' @usage - DivCell range�� getter
Public Property Get DivCell() As Range
    Set DivCell = pDivCell
End Property

' Setter for DivCell Range
' @method DivCell
' @param {Range} value
' @usage - DivCell range�� setter
Public Property Set DivCell(value As Range)
    Set pDivCell = value
End Property

' Getter for DataIdRange
' @method DataIdRange
' @return {Range}
' @usage - DataIdRange�� getter
Public Property Get DataIdRange() As Range
    Set DataIdRange = pdataIdRange
End Property

' Setter for DataIdRange
' @method DataIdRange
' @param {Range} value
' @usage - DataIdRange�� setter
Public Property Set DataIdRange(value As Range)
    Set pdataIdRange = value
End Property

' jsonConverter library�� ���� reference�� �ʿ�� �Ѵ�.
' jsonConverter.bas file�� project�� import�ؾ� �Ѵ�.
Public Function GenerateJson2() As String
    Dim dataList As Collection
    Set dataList = New Collection
    
    Dim cell As Range
    Dim Item As Dictionary
    Dim dataId As String
    Dim yield As Double
    
    For Each cell In pdataIdRange
        Set Item = New Dictionary
        dataId = cell.value
        yield = cell.Offset(0, 1).value
        
        With Item
            .Add "dataId", dataId
            .Add "yield", yield
        End With
        
        dataList.Add Item
        DoEvents
    Next cell
    
    ' jsonConverter�� ����ؼ� JSON string���� collection�� �����Ѵ�.
    GenerateJson2 = JsonConverter.ConvertToJson(dataList)
End Function
