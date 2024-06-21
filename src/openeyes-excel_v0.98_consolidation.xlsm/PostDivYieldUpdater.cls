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
' @usage - worksheet의 getter
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWs
End Property

' Setter for Worksheet
' @method Worksheet
' @param {Worksheet} value
' @usage - worksheet의 setter
Public Property Set Worksheet(value As Worksheet)
    Set pWs = value
End Property

' Getter for DivCell Range
' @method DivCell
' @return {Range}
' @usage - DivCell range의 getter
Public Property Get divCell() As Range
    Set divCell = pDivCell
End Property

' Setter for DivCell Range
' @method DivCell
' @param {Range} value
' @usage - DivCell range의 setter
Public Property Set divCell(value As Range)
    Set pDivCell = value
End Property

' Getter for DataIdRange
' @method DataIdRange
' @return {Range}
' @usage - DataIdRange의 getter
Public Property Get dataIdRange() As Range
    Set dataIdRange = pdataIdRange
End Property

' Setter for DataIdRange
' @method DataIdRange
' @param {Range} value
' @usage - DataIdRange의 setter
Public Property Set dataIdRange(value As Range)
    Set pdataIdRange = value
End Property

' jsonConverter library에 대한 reference를 필요로 한다.
' jsonConverter.bas file을 project에 import해야 한다.
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
    
    ' jsonConverter를 사용해서 JSON string으로 collection을 변경한다.
    GenerateJson2 = JsonConverter.ConvertToJson(dataList)
End Function

