Attribute VB_Name = "PostVol2"
Option Explicit

Function ConvertRangeToJSON() As String
    Dim ws As Worksheet
    Dim cell As Range
    Dim jsonString As String
    Dim DataId As String
    Dim firstObject As Boolean
    
    Set ws = ThisWorkbook.Sheets("Vol")
    jsonString = "["
    
    firstObject = True
    For Each cell In ws.Range("AD1:AD" & ws.Cells(ws.Rows.Count, "AD").End(xlUp).Row)
        Select Case cell.Value
            Case "KOSPI_LV"
                DataId = "KOSPI200_LOC"
            Case "NKY_LV"
                DataId = "N225_LOC"
            Case "HSI_LV"
                DataId = "HSI_LOC"
            Case "HSCEI_LV"
                DataId = "HSCEI_LOC"
            Case Else
                DataId = "" '값이 어떤 case에도 맞지 않다면 skip한다.
        End Select
        
        If DataId <> "" Then
            If Not firstObject Then
                jsonString = jsonString & ","
            End If
            jsonString = jsonString & GenerateObjectJSON(ws, cell, DataId)
            firstObject = False
        End If
    Next cell
    
    jsonString = jsonString & "]"
    ConvertRangeToJSON = jsonString
End Function

Function GenerateObjectJSON(ws As Worksheet, RefCell As Range, DataId As String) As String
    Dim volFactorRange As Range, tenorRange As Range, dataRange As Range
    Dim volFactorCell As Range, termVolCell As Range
    Dim firstTermVol As Boolean, firstVolCurve As Boolean
    Dim objectJSON As String
    
    ' refCell의 데이터와 volFactor, tenor에 근거해서 range를 잡는다.
    Set volFactorRange = ws.Range(RefCell.Offset(0, 2), RefCell.Offset(0, 2).End(xlToRight))
    Set tenorRange = ws.Range(RefCell.Offset(1, 1), RefCell.Offset(1, 1).End(xlDown))
    Set dataRange = ws.Range(volFactorRange.Offset(1, 0), tenorRange.Offset(0, volFactorRange.Columns.Count - 1))
    
    objectJSON = "{" & """dataId"": """ & DataId & """," & """volCurves"": ["
    
    firstVolCurve = True
    For Each volFactorCell In volFactorRange
        If Not firstVolCurve Then
            objectJSON = objectJSON & ","
        End If
        objectJSON = objectJSON & "{" & """termVols"": ["
        
        firstTermVol = True
        For Each termVolCell In tenorRange
            If Not firstTermVol Then
                objectJSON = objectJSON & ","
            End If
            objectJSON = objectJSON & "{" & """tenor"": " & termVolCell.Value & "," & """vol"": " & ws.Cells(termVolCell.Row, volFactorCell.Column).Value & "}"
            firstTermVol = False
        Next termVolCell
        
        objectJSON = objectJSON & "]," & """volFactor"": " & volFactorCell.Value & "}"
        firstVolCurve = False
    Next volFactorCell
    
    objectJSON = objectJSON & "]}"
    
    GenerateObjectJSON = objectJSON
End Function

Sub RunFunc()
    Dim jsonString As String
    jsonString = ConvertRangeToJSON()
    'Debug.Print jsonString
    jsonString = URLEncode(jsonString)
    
    
    ' request를 보낼 URL
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/vols?baseDt=20231228&dataSetId=TEST11"
    
    SendPostRequest jsonString, url

End Sub

