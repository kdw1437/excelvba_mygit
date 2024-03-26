Attribute VB_Name = "PostVol2"
Option Explicit

Function ConvertRangeToJSON() As String
    Dim Ws As Worksheet
    Dim cell As Range
    Dim jsonString As String
    Dim dataId As String
    Dim firstObject As Boolean
    
    Set Ws = ThisWorkbook.Sheets("Vol")
    jsonString = "["
    
    firstObject = True
    For Each cell In Ws.Range("AD1:AD" & Ws.Cells(Ws.Rows.Count, "AD").End(xlUp).Row)
        Select Case cell.Value
            Case "KOSPI_LV"
                dataId = "KOSPI200_LOC"
            Case "NKY_LV"
                dataId = "N225_LOC"
            Case "HSI_LV"
                dataId = "HSI_LOC"
            Case "HSCEI_LV"
                dataId = "HSCEI_LOC"
            Case Else
                dataId = "" '���� � case���� ���� �ʴٸ� skip�Ѵ�.
        End Select
        
        If dataId <> "" Then
            If Not firstObject Then
                jsonString = jsonString & ","
            End If
            jsonString = jsonString & GenerateObjectJSON(Ws, cell, dataId)
            firstObject = False
        End If
    Next cell
    
    jsonString = jsonString & "]"
    ConvertRangeToJSON = jsonString
End Function

Function GenerateObjectJSON(Ws As Worksheet, RefCell As Range, dataId As String) As String
    Dim volFactorRange As Range, tenorRange As Range, dataRange As Range
    Dim volFactorCell As Range, termVolCell As Range
    Dim firstTermVol As Boolean, firstVolCurve As Boolean
    Dim objectJSON As String
    
    ' refCell�� �����Ϳ� volFactor, tenor�� �ٰ��ؼ� range�� ��´�.
    Set volFactorRange = Ws.Range(RefCell.Offset(0, 2), RefCell.Offset(0, 2).End(xlToRight))
    Set tenorRange = Ws.Range(RefCell.Offset(1, 1), RefCell.Offset(1, 1).End(xlDown))
    Set dataRange = Ws.Range(volFactorRange.Offset(1, 0), tenorRange.Offset(0, volFactorRange.Columns.Count - 1))
    
    objectJSON = "{" & """dataId"": """ & dataId & """," & """volCurves"": ["
    
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
            objectJSON = objectJSON & "{" & """tenor"": " & termVolCell.Value & "," & """vol"": " & Ws.Cells(termVolCell.Row, volFactorCell.Column).Value & "}"
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
    Debug.Print jsonString
    jsonString = URLEncode(jsonString)
    
    
    ' request�� ���� URL
    Dim url As String
    url = "http://localhost:8080/val/marketdata/v1/vols?baseDt=20231228&dataSetId=TEST11"
    
    SendPostRequest jsonString, url

End Sub

