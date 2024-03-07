Attribute VB_Name = "Volatility"
Sub vol()


    Dim VoUrlBuilder As UrlBuilder
    Set VoUrlBuilder = New UrlBuilder
    
    VoUrlBuilder.baseURL = "http://localhost:8080/val/marketdata/"
    VoUrlBuilder.Version = "v1/"
    VoUrlBuilder.DataParameter = "vols?"
    VoUrlBuilder.baseDt = "baseDt=20231228&"
    VoUrlBuilder.DataIds = "dataIds=HSCEI_LOC,HSI_LOC,N225_LOC,KOSPI200_LOC"
    
    Dim VoUrl As String
    VoUrl = VoUrlBuilder.MakeUrl
    
    Debug.Print VoUrl
    
    Dim JsonResponse As Object
    Set JsonResponse = GetJsonResponse(VoUrl)
    Dim volatilities As Collection
    Set volatilities = JsonResponse("response")("volatilities")
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Vol")
    
    Dim volCurve As Variant
    Dim termVol As Variant
    Dim dataId As String
    Dim code As String
    Dim r As Long, c As Long
    
    For Each volCurve In volatilities
        dataId = volCurve("dataId")
        code = MapDataIdToCode(dataId)
        
        Dim codeRow As Range
        Set codeRow = ws.Columns("A:A").Find(What:=code, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not codeRow Is Nothing Then
            Dim codeRowNumber As Long
            codeRowNumber = codeRow.Row
            Dim searchRange1 As Range
            Set searchRange1 = ws.Range(ws.Cells(codeRowNumber, 3), ws.Cells(codeRowNumber, 3).End(xlToRight))
            Dim searchRange2 As Range
            'Excel에서 연속된 sheet를 지정할 때, xlDown과 xlRight를 주로 사용한다. 중간에 빈셀이 있을 시, 그걸 고려해서 그 앞 지점에서 범위 지정이 멈추게 된다.
            Set searchRange2 = ws.Range(ws.Cells(codeRowNumber + 1, 2), ws.Cells(codeRowNumber + 1, 2).End(xlDown))
            For Each termVol In volCurve("volCurves")
                Dim volFactor As Double
                volFactor = termVol("volFactor")

                
                c = ws.Rows(codeRowNumber).Find(What:=volFactor, LookIn:=xlValues, LookAt:=xlWhole).Column
                For Each volEntry In termVol("termVols")
                    Dim tenor As Double
                    tenor = volEntry("tenor")
                    
                    
                    
                    Dim tenorCell As Range
                    Set tenorCell = searchRange2.Find(What:=tenor, LookIn:=xlValues, LookAt:=xlWhole)
                    
                    If Not tenorCell Is Nothing Then
                        r = tenorCell.Row
                        ws.Cells(r, c).value = volEntry("vol")
                    End If
                Next volEntry
            Next termVol
        Dim headerCell As Range
        Dim rowHeaderCell As Range
        Dim dataCell As Range
        
        For Each headerCell In searchRange1
            For Each rowHeaderCell In searchRange2
                Set dataCell = ws.Cells(rowHeaderCell.Row, headerCell.Column)
                If IsEmpty(dataCell.value) Then
                    dataCell.value = 0
                End If
            Next rowHeaderCell
        Next headerCell
        End If
    Next volCurve
End Sub

Function MapDataIdToCode(dataId As String) As String
    ' mapping함수
    Select Case dataId
        Case "KOSPI200_LOC"
            MapDataIdToCode = "KOSPI_LV"
        Case "HSI_LOC"
            MapDataIdToCode = "HSI_LV"
        Case "N225_LOC"
            MapDataIdToCode = "NKY_LV"
        Case "HSCEI_LOC"
            MapDataIdToCode = "HSCEI_LV"
        
        Case Else
            MapDataIdToCode = dataId ' mapping에 해당되는 경우가 없을 때는, 그대로 return한다.
    End Select
End Function
