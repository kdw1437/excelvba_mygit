Attribute VB_Name = "PostVol"
Option Explicit

Function ConvertRangeToJson() As String
    Dim ws As Worksheet
    Dim volCurve As Object
    Dim termVols As Object
    Dim volFactorRange As Range, tenorRange As Range, dataRange As Range
    Dim i As Long, j As Long
    'Dim volFactor As Double, tenor As Double, vol As Double
    Dim dataId As String
    Dim jsonString As String
    
    
    
    Set ws = ThisWorkbook.Sheets("Vol")
    
    jsonString = "["
    
    Dim RefCell As Range
    Dim cell As Range
    Dim found As Boolean
    
    For Each cell In ws.Range("AD:AD")
        If cell.value = "KOSPI_LV" Then
            Set RefCell = cell
            found = True
            Exit For
        End If
    Next cell
    
    ' Set the volFactorRange by starting at AF4 and going to the right till the last contiguous cell
    Set volFactorRange = ws.Range(RefCell.Offset(0, 2), RefCell.Offset(0, 2).End(xlToRight))
    
    ' Set the tenorRange by starting at AE5 and going down till the last contiguous cell
    Set tenorRange = ws.Range(RefCell.Offset(1, 1), RefCell.Offset(1, 1).End(xlDown))
    
    ' Set the dataRange based on the volFactorRange and tenorRange
    Set dataRange = ws.Range(volFactorRange.Offset(1, 0), tenorRange.Offset(0, volFactorRange.Columns.Count - 1))
    
    jsonString = jsonString & "{" & """dataId"": ""KOSPI200_LOC""," & """volCurves"": ["
    
    
    Dim volFactorCell As Range
    Dim termVolCell As Range
    
    Dim volFactor As Variant
    Dim tenor As Variant
    Dim vol As Variant
    Dim firstTermVol As Boolean
    Dim firstVolCurve As Boolean
        
    firstVolCurve = True
    For Each volFactorCell In volFactorRange
        volFactor = volFactorCell.value
        If Not firstVolCurve Then
            jsonString = jsonString & ","
        End If
        jsonString = jsonString & "{" & """termVols"": ["
        
        firstTermVol = True
        For Each termVolCell In tenorRange
            tenor = termVolCell.value
            vol = ws.Cells(termVolCell.row, volFactorCell.Column).value
            If Not firstTermVol Then
                jsonString = jsonString & ","
            End If
            jsonString = jsonString & "{" & """tenor"": " & tenor & "," & """vol"": " & vol & "}"
            
            firstTermVol = False
        Next termVolCell
        
        jsonString = jsonString & "]," & """volFactor"": " & volFactor & "}"
        
        firstVolCurve = False
    Next volFactorCell
    
    jsonString = jsonString & "]" & "}"
    
    ' Close JSON array
    jsonString = jsonString & "]"
    Debug.Print jsonString
' Now you have volFactorRange, tenorRange, and dataRange as per the specified cells
    ConvertRangeToJson = jsonString
End Function


Sub RunFunc()
    Dim jsonString As String
    jsonString = ConvertRangeToJson()
End Sub

