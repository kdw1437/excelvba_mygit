Attribute VB_Name = "postDisDivModule"
Option Explicit

Sub postDisDivModule()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Missing Data - D_Dividend")
    
    Dim StartCell As Range
    'Set startCell = ws.Cells(3, 1)
    Set StartCell = ws.Range("A:A").Find(What:="Discrete Dividend", Lookat:=xlWhole)
'    If startCell Is Nothing Then
'        MsgBox "Discrete Dividend not found."
'        Exit Sub
'    End If

    
    Dim dividendStreamCollection As New Collection '가장 바깥의 array는 Collection으로 선언한다.
    Dim i As Long, j As Long, k As Long
    k = 4  ' This should be dynamically determined if possible
    
    ' Loop through each currency block and collect the data
    For i = 1 To k
        Dim dataIdCell As Range
        
        'Debug.Print "Trying to access: " & startCell.Offset(3, 1 + 3 * (i - 1)).Address
        
        'dataIdCell = startCell.Offset(3, 1 + 3 * (i - 1))
        'On Error Resume Next
        Set dataIdCell = StartCell.Offset(3, 1 + 3 * (i - 1))
'        If Err.Number <> 0 Then
'            MsgBox "Failed to set dataIdCell: " & Err.Description
'            Err.Clear
'            Exit Sub
'        ElseIf dataIdCell Is Nothing Then
'            MsgBox "dataIdCell is not set, possibly invalid offset."
'            Exit Sub
'        End If
'        On Error GoTo 0

        Dim dataId As String
        dataId = dataIdCell.value
        
        Dim disDivData As Object
        Set disDivData = CreateObject("Scripting.Dictionary") 'array안의 object는 object선언 뒤, Dictionary instance를 받는다.
       
        disDivData("dataId") = dataId
        
        Dim dividendsCollection As New Collection 'dividendsCollection array선언. 안에 dictionary의 array를 value로 받는다.
        Set dividendsCollection = New Collection 'Reinitialize
        
        j = 4 ' Start row offset for yields
        Do While ws.Cells(dataIdCell.Row + j, dataIdCell.Column).value <> ""
            Dim divData As Object
            Set divData = CreateObject("Scripting.Dictionary")
            
            divData("date") = ws.Cells(dataIdCell.Row + j, dataIdCell.Column - 1).value
            divData("value") = ws.Cells(dataIdCell.Row + j, dataIdCell.Column).value
            
            dividendsCollection.Add divData
            j = j + 1
        Loop
        'Collection은 dictionary의 key의 value가 될 수 없다.
        'collection을 array로 바꿔 주고, array를 value로 key와 함께 할당한다.
        Dim divsArray() As Variant
        ReDim divsArray(1 To dividendsCollection.Count)
        
        Dim idx As Long
        For idx = 1 To dividendsCollection.Count
            Set divsArray(idx) = dividendsCollection(idx)
        Next idx
        
        '이제 array를 dictionary에 할당한다.
        disDivData("dividends") = divsArray
        dividendStreamCollection.Add disDivData
    Next i
    
    ' 전체 collection을 JSON으로 변경한다.
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(dividendStreamCollection)
    
    Debug.Print jsonString
End Sub

Sub UseDividendDataProcessor()
    Dim dp As postDisDiv
    Set dp = New postDisDiv
    
    Set dp.Worksheet = ThisWorkbook.Worksheets("Missing Data - D_Dividend")
    Set dp.StartCell = dp.Worksheet.Range("A:A").Find(What:="Discrete Dividend", Lookat:=xlWhole)
    dp.k = 4 ' (K가 알려진 값이고 dynamic하게 결정되어질 때는 거기에 맞춰서 코드 작성)
    
    If Not dp.StartCell Is Nothing Then
        Dim jsonString As String
        
        jsonString = dp.ReturnJSON
    Else
        MsgBox "Start cell not found."
    End If
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveDividendStream?baseDt=20240412&dataSetId=official"
    
    ' JSON data와 POST request를 보내는 subroutine을 호출한다.
    SendPostRequest jsonString, url
    
End Sub

