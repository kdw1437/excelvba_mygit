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

    
    Dim dividendStreamCollection As New Collection '���� �ٱ��� array�� Collection���� �����Ѵ�.
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
        Set disDivData = CreateObject("Scripting.Dictionary") 'array���� object�� object���� ��, Dictionary instance�� �޴´�.
       
        disDivData("dataId") = dataId
        
        Dim dividendsCollection As New Collection 'dividendsCollection array����. �ȿ� dictionary�� array�� value�� �޴´�.
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
        'Collection�� dictionary�� key�� value�� �� �� ����.
        'collection�� array�� �ٲ� �ְ�, array�� value�� key�� �Բ� �Ҵ��Ѵ�.
        Dim divsArray() As Variant
        ReDim divsArray(1 To dividendsCollection.Count)
        
        Dim idx As Long
        For idx = 1 To dividendsCollection.Count
            Set divsArray(idx) = dividendsCollection(idx)
        Next idx
        
        '���� array�� dictionary�� �Ҵ��Ѵ�.
        disDivData("dividends") = divsArray
        dividendStreamCollection.Add disDivData
    Next i
    
    ' ��ü collection�� JSON���� �����Ѵ�.
    Dim jsonString As String
    jsonString = JsonConverter.ConvertToJson(dividendStreamCollection)
    
    Debug.Print jsonString
End Sub

Sub UseDividendDataProcessor()
    Dim dp As postDisDiv
    Set dp = New postDisDiv
    
    Set dp.Worksheet = ThisWorkbook.Worksheets("Missing Data - D_Dividend")
    Set dp.StartCell = dp.Worksheet.Range("A:A").Find(What:="Discrete Dividend", Lookat:=xlWhole)
    dp.k = 4 ' (K�� �˷��� ���̰� dynamic�ϰ� �����Ǿ��� ���� �ű⿡ ���缭 �ڵ� �ۼ�)
    
    If Not dp.StartCell Is Nothing Then
        Dim jsonString As String
        
        jsonString = dp.ReturnJSON
    Else
        MsgBox "Start cell not found."
    End If
    
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/saveDividendStream?baseDt=20240412&dataSetId=official"
    
    ' JSON data�� POST request�� ������ subroutine�� ȣ���Ѵ�.
    SendPostRequest jsonString, url
    
End Sub

