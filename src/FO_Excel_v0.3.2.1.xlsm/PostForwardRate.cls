VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostForwardRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pWorksheet As Worksheet
Private pStartCell As Range
Private pK As Long

' Property for the worksheet
'
' @property Get Worksheet
' @return {Worksheet} worksheet 객체
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWorksheet
End Property

' @property Set Worksheet
' @param {Worksheet} value - worksheet object를 set한다.
Public Property Set Worksheet(value As Worksheet)
    Set pWorksheet = value
End Property

' Property for the start cell
'
' @property Get StartCell
' @return {Range} 시작 Cell range
Public Property Get StartCell() As Range
    Set StartCell = pStartCell
End Property

' @property Set StartCell
' @param {Range} value - 시작 cell range를 set한다.
Public Property Set StartCell(value As Range)
    Set pStartCell = value
End Property

' Property for the number of currencies
'
' @property Get k
' @return {Long} k값
Public Property Get k() As Long
    k = pK
End Property

' @property Let k
' @param {Long} value - k값을 set한다.
Public Property Let k(value As Long)
    pK = value
End Property

' JSON String을 생성해서 return하는 메소드
'
' @method ReturnJSON
' @return {String} worksheet의 data로 부터 만들어진 JSON string
Public Function ReturnJSON() As String
    Dim currenciesCollection As New Collection
    Dim i As Long, j As Long
    
    For i = 1 To pK ' Use the K property here
        Dim relCurrencyCell As Range
        Dim currencyCode As String
        currencyCode = pStartCell.Offset(4, 1 + 3 * (i - 1)).value
        
        Set relCurrencyCell = pWorksheet.Rows(7).Find(What:=currencyCode, Lookat:=xlWhole)
        
        Dim currencyData As Object
        Set currencyData = CreateObject("Scripting.Dictionary")
        
        Dim dataId As String
        dataId = "F_FX" & currencyCode & pStartCell.Offset(3, 1 + 3 * (i - 1)).value
        currencyData("dataId") = dataId
        
        Dim yieldsCollection As New Collection
        Set yieldsCollection = New Collection '재초기화
        
        j = 2 ' Start row offset for yields
        Do While pWorksheet.Cells(relCurrencyCell.Row + j, relCurrencyCell.Column).value <> ""
            Dim yieldData As Object
            Set yieldData = CreateObject("Scripting.Dictionary")
            yieldData("tenor") = pWorksheet.Cells(relCurrencyCell.Row + j, relCurrencyCell.Column - 1).value
            yieldData("value") = pWorksheet.Cells(relCurrencyCell.Row + j, relCurrencyCell.Column).value
            yieldsCollection.Add yieldData
            j = j + 1
        Loop
        
        Dim yieldsArray() As Variant
        Dim idx As Long
        
        ReDim yieldsArray(1 To yieldsCollection.Count)
        For idx = 1 To yieldsCollection.Count
            Set yieldsArray(idx) = yieldsCollection(idx)
        Next idx
        
        currencyData("yields") = yieldsArray
        currenciesCollection.Add currencyData
    Next i
    
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(currenciesCollection)
    
    Debug.Print jsonString
    
    ReturnJSON = jsonString
End Function


