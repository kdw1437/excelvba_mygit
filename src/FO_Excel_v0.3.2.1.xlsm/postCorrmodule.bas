Attribute VB_Name = "postCorrmodule"
Option Explicit

Sub postCorrModule()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Missing Data - Hist Vol, Corr")
    
    Dim jsonString As String
    jsonString = "["
    
    Dim i As Long
    i = 5
    
    ' Loop until an empty cell is found in column E
    Do While ws.Cells(i, "E").value <> ""
        ' Construct the JSON string
        jsonString = jsonString & "{""dataId"":""" & ws.Cells(i, "E").value & ":" & ws.Cells(i, "F").value & ""","
        jsonString = jsonString & """dataId1"":""" & ws.Cells(i, "E").value & ""","
        jsonString = jsonString & """dataId2"":""" & ws.Cells(i, "F").value & ""","
        jsonString = jsonString & """corr"":" & ws.Cells(i, "G").value & "}"
        
        i = i + 1 ' Move to the next row
        
        ' If the next pair of dataId1 or dataId2 is not empty, add a comma to separate the objects
        If ws.Cells(i, "E").value <> "" Or ws.Cells(i, "F").value <> "" Then
            jsonString = jsonString & ","
        End If
    Loop
    
    jsonString = jsonString & "]"
    
    ' Output the jsonString to Immediate Window (Ctrl + G in VBA Editor to view)
    Debug.Print jsonString
    

End Sub

Sub postCorrModule2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Missing Data - Hist Vol, Corr")
    
    Dim dataCollection As New Collection
    Dim i As Long
    i = 5
    
    ' 칼럼 E에서 빈 cell이 발견될 때 까지 Loop를 돈다.
    Do While ws.Cells(i, "E").value <> ""
        Dim dataDict As Object
        Set dataDict = CreateObject("Scripting.Dictionary")
        
        ' 만들어진 dictionary를 cell에 있는 데이터로 key, value 형태로 채운다. (jsonObject를 생성하는 과정이다.)
        dataDict.Add "dataId", ws.Cells(i, "E").value & ":" & ws.Cells(i, "F").value
        dataDict.Add "dataId1", ws.Cells(i, "E").value
        dataDict.Add "dataId2", ws.Cells(i, "F").value
        dataDict.Add "corr", ws.Cells(i, "G").value
        
        ' Collection에 생성된 dictionary를 넣는다. (collection은 json array, dictionary는 json object이다.)
        dataCollection.Add dataDict
        
        i = i + 1 ' 다음 row로 이동한다.
    Loop
    
    ' JsonConverter의 ConvertToJson메소드를 이용해서 collection(dictionary의 collection을 json object의 json array로 바꾼다.)
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(dataCollection, Whitespace:=2)
    
    ' Console에 출력
    Debug.Print jsonString
End Sub

Sub UseCorrelationDataProcessor()
    Dim corrProcessor As New postCorr
    
    ' Set properties
    Set corrProcessor.Worksheet = ThisWorkbook.Sheets("Missing Data - Hist Vol, Corr")
    corrProcessor.StartRow = 5
    corrProcessor.Column = "E"
    
    ' Generate JSON
    Dim jsonString As String
    jsonString = corrProcessor.GenerateJSON
    
    ' Print JSON
    Debug.Print jsonString
End Sub

