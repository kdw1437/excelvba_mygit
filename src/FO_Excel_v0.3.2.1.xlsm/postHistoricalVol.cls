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
Public Property Get StartRow() As Long
    StartRow = pStartRow
End Property

Public Property Let StartRow(value As Long)
    pStartRow = value
End Property

Public Sub ProcessData()
    ' Jsonarray�� �ٲ� CollectionŸ�� dataCollection
    Dim dataCollection As New Collection
    
    Dim i As Long
    i = pStartRow
    
    ' �� currency block�� Loop���鼭 data�� ������.
    Do While pWorksheet.Cells(i, "A").value <> ""
        Dim dataDict As Object
        Set dataDict = CreateObject("Scripting.Dictionary")
        
        ' Dictionary�� key value ���� ���� �ִ´�.
        dataDict.Add "dataId", pWorksheet.Cells(i, "A").value & "_VOL_250"
        dataDict.Add "histvol", pWorksheet.Cells(i, "C").value / 100
        
        ' Collection(json array)�� Dictionary(json object)�� ���� �ִ´�.
        dataCollection.Add dataDict
        
        i = i + 1
    Loop
    
    ' Collection type data�� json array�� �ٲ��ش�. (JsonConverter�� ConvertToJson�̿�)
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(dataCollection, Whitespace:=2)
    
    ' JSON string ���
    Debug.Print jsonString
End Sub
