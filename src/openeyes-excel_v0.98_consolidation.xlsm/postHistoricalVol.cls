VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "postHistoricalVol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pWorksheet As Worksheet
Private pStartRow As Long

' Property for the worksheet
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWorksheet
End Property

Public Property Set Worksheet(value As Worksheet)
    Set pWorksheet = value
End Property

' Property for the starting row
Public Property Get startRow() As Long
    startRow = pStartRow
End Property

Public Property Let startRow(value As Long)
    pStartRow = value
End Property

Public Sub ProcessData()
    ' Jsonarray로 바뀔 Collection타입 dataCollection
    Dim dataCollection As New Collection
    
    Dim i As Long
    i = pStartRow
    
    ' 각 currency block을 Loop돌면서 data를 모은다.
    Do While pWorksheet.Cells(i, "A").value <> ""
        Dim dataDict As Object
        Set dataDict = CreateObject("Scripting.Dictionary")
        
        ' Dictionary에 key value 쌍을 집어 넣는다.
        dataDict.Add "dataId", pWorksheet.Cells(i, "A").value & "_VOL_250"
        dataDict.Add "historicalVol", pWorksheet.Cells(i, "C").value / 100
        
        ' Collection(json array)에 Dictionary(json object)를 집어 넣는다.
        dataCollection.Add dataDict
        
        i = i + 1
    Loop
    
    ' Collection type data를 json array로 바꿔준다. (JsonConverter의 ConvertToJson이용)
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(dataCollection, Whitespace:=2)
    
    ' JSON string 출력
    Debug.Print jsonString
End Sub

Public Function ReturnData() As String
    ' Jsonarray로 바뀔 Collection타입 dataCollection
    Dim dataCollection As New Collection
    
    Dim i As Long
    i = pStartRow
    
    ' 각 currency block을 Loop돌면서 data를 모은다.
    Do While pWorksheet.Cells(i, "A").value <> ""
        Dim dataDict As Object
        Set dataDict = CreateObject("Scripting.Dictionary")
        
        ' Dictionary에 key value 쌍을 집어 넣는다.
        dataDict.Add "dataId", pWorksheet.Cells(i, "A").value & "_VOL_250"
        dataDict.Add "historicalVol", pWorksheet.Cells(i, "C").value / 100
        
        ' Collection(json array)에 Dictionary(json object)를 집어 넣는다.
        dataCollection.Add dataDict
        
        i = i + 1
    Loop
    
    ' Collection type data를 json array로 바꿔준다. (JsonConverter의 ConvertToJson이용)
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(dataCollection, Whitespace:=2)
    
    ' JSON string 출력
    Debug.Print jsonString
    
    ReturnData = jsonString
End Function


