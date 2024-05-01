VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "postDisDiv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Class Module: postDisDiv
Private pWorksheet As Worksheet
Private pStartCell As Range
Private pK As Long

' Property for Worksheet
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWorksheet
End Property
Public Property Set Worksheet(value As Worksheet)
    Set pWorksheet = value
End Property

' Property for StartCell
Public Property Get StartCell() As Range
    Set StartCell = pStartCell
End Property
Public Property Set StartCell(value As Range)
    Set pStartCell = value
End Property

' Property for K
Public Property Get k() As Long
    k = pK
End Property
Public Property Let k(value As Long)
    pK = value
End Property

' Method to Generate JSON
Public Sub GenerateJSON()
    Dim dividendStreamCollection As New Collection '가장 바깥의 jsonArray
    Dim i As Long, j As Long

    For i = 1 To pK
        Dim dataIdCell As Range
        Set dataIdCell = pStartCell.Offset(3, 1 + 3 * (i - 1))
        
        Dim dataId As String
        dataId = dataIdCell.value
        
        Dim disDivData As Object
        Set disDivData = CreateObject("Scripting.Dictionary")
        disDivData("dataId") = dataId '가장 바깥의 jsonArray의 jsonObject의 key dataId와 value값 할당
        
        Dim dividendsCollection As New Collection 'jsonObject의 key의 value값이 될 jsonarray 생성
        Set dividendsCollection = New Collection '재초기화 (jsonArray 재초기화)
        j = 4 ' Start row offset for yields
        
        Do While pWorksheet.Cells(dataIdCell.Row + j, dataIdCell.Column).value <> ""
            Dim divData As Object
            Set divData = CreateObject("Scripting.Dictionary") 'jsonarray의 요소가 될 jsonObject 생성
            divData("date") = pWorksheet.Cells(dataIdCell.Row + j, dataIdCell.Column - 1).value
            divData("value") = pWorksheet.Cells(dataIdCell.Row + j, dataIdCell.Column).value
            dividendsCollection.Add divData 'divData jsonObject(Dictionary)를 jsonArray에 넣어준다.
            j = j + 1
        Loop
        
        Dim divsArray() As Variant
        ReDim divsArray(1 To dividendsCollection.Count) '만들어진 dictionary를 요소로가지는 collection을 Array형태로 바꿔서
        'structured된 json객체를 구현한다.
        
        Dim idx As Long
        For idx = 1 To dividendsCollection.Count
            Set divsArray(idx) = dividendsCollection(idx) 'array의 첫번째에 collection의 첫번째 dictionary를 넣어준다.
        Next idx
        
        disDivData("dividends") = divsArray 'dictionary의 key dividends에 divsArray(json array) value를 할당한다.
        dividendStreamCollection.Add disDivData '모든 값이 할당된 dictionary를 collection에 순서대로 추가한다.
    Next i
    
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(dividendStreamCollection) 'collection을 ConvertToJson메소드를 이용해서 jsonString으로 변환한다.
    
    Debug.Print jsonString
End Sub

' Method to Generate JSON
Public Function ReturnJSON() As String
    Dim dividendStreamCollection As New Collection '가장 바깥의 jsonArray
    Dim i As Long, j As Long

    For i = 1 To pK
        Dim dataIdCell As Range
        Set dataIdCell = pStartCell.Offset(3, 1 + 3 * (i - 1))
        
        Dim dataId As String
        dataId = dataIdCell.value
        
        Dim disDivData As Object
        Set disDivData = CreateObject("Scripting.Dictionary")
        disDivData("dataId") = dataId '가장 바깥의 jsonArray의 jsonObject의 key dataId와 value값 할당
        
        Dim dividendsCollection As New Collection 'jsonObject의 key의 value값이 될 jsonarray 생성
        Set dividendsCollection = New Collection '재초기화 (jsonArray 재초기화)
        j = 4 ' Start row offset for yields
        
        Do While pWorksheet.Cells(dataIdCell.Row + j, dataIdCell.Column).value <> ""
            Dim divData As Object
            Set divData = CreateObject("Scripting.Dictionary") 'jsonarray의 요소가 될 jsonObject 생성
            divData("date") = CStr(pWorksheet.Cells(dataIdCell.Row + j, dataIdCell.Column - 1).value)
            divData("value") = pWorksheet.Cells(dataIdCell.Row + j, dataIdCell.Column).value
            dividendsCollection.Add divData 'divData jsonObject(Dictionary)를 jsonArray에 넣어준다.
            j = j + 1
        Loop
        
        Dim divsArray() As Variant
        ReDim divsArray(1 To dividendsCollection.Count) '만들어진 dictionary를 요소로가지는 collection을 Array형태로 바꿔서
        'structured된 json객체를 구현한다.
        
        Dim idx As Long
        For idx = 1 To dividendsCollection.Count
            Set divsArray(idx) = dividendsCollection(idx) 'array의 첫번째에 collection의 첫번째 dictionary를 넣어준다.
        Next idx
        
        disDivData("dividends") = divsArray 'dictionary의 key dividends에 divsArray(json array) value를 할당한다.
        dividendStreamCollection.Add disDivData '모든 값이 할당된 dictionary를 collection에 순서대로 추가한다.
    Next i
    
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(dividendStreamCollection) 'collection을 ConvertToJson메소드를 이용해서 jsonString으로 변환한다.
    
    Debug.Print jsonString
    ReturnJSON = jsonString
    
End Function


