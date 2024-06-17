VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostCorrUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: PostCorrUpdater
' 상관 데이터를 저장하는 class module. 두 개의 range를 받아 각각의 데이터 쌍에 대한 상관 값을 JSON 형태로 반환한다.

'Private 변수 선언
Private pVerticalRange As Range
Private pHorizontalRange As Range

' vertical range setter
'@method VerticalRange
'@param {Range} rng
'@usage - Vertical range를 설정하는 property 메서드
Public Property Set VerticalRange(rng As Range)
    Set pVerticalRange = rng
End Property

' horizontal range setter
'@method HorizontalRange
'@param {Range} rng
'@usage - Horizontal range를 설정하는 property 메서드
Public Property Set HorizontalRange(rng As Range)
    Set pHorizontalRange = rng
End Property

'@method CorrJsonh2
'@return {String}
'@usage - Horizontal 데이터를 기준으로 상관 데이터를 JSON 형식으로 반환하는 함수
Public Function CorrJsonh2() As String
    Dim correlationData As New Collection
    Dim vCell As Range, hCell As Range
    Dim pairData As Dictionary
    Dim uniquePairs As New Scripting.Dictionary

    For Each vCell In pVerticalRange
        For Each hCell In pHorizontalRange
            Dim dataId1 As String
            Dim dataId2 As String
            Dim corrValue As Variant

            dataId1 = hCell.value
            dataId2 = vCell.value
            corrValue = vCell.Worksheet.Cells(vCell.row, hCell.Column).value

            If Not IsEmpty(corrValue) Then
                If Not IsDuplicate(dataId1, dataId2, uniquePairs) Then
                    Set pairData = New Dictionary
                    With pairData
                        .Add "dataId1", dataId1
                        .Add "dataId2", dataId2
                        .Add "dataId", dataId1 & ":" & dataId2
                        .Add "corr", corrValue
                    End With
                    correlationData.Add pairData
                    uniquePairs.Add dataId1 & ":" & dataId2, Nothing
                End If
            End If
            DoEvents
        Next hCell
        DoEvents
    Next vCell

    CorrJsonh2 = JsonConverter.ConvertToJson(correlationData)
End Function

'@method CorrJsonv2
'@return {String}
'@usage - Vertical 데이터를 기준으로 상관 데이터를 JSON 형식으로 반환하는 함수
Public Function CorrJsonv2() As String
    Dim correlationData As New Collection
    Dim vCell As Range, hCell As Range
    Dim pairData As Dictionary
    Dim uniquePairs As New Scripting.Dictionary

    For Each vCell In pVerticalRange
        For Each hCell In pHorizontalRange
            Dim dataId1 As String
            Dim dataId2 As String
            Dim corrValue As Variant

            dataId1 = vCell.value
            dataId2 = hCell.value
            corrValue = vCell.Worksheet.Cells(vCell.row, hCell.Column).value

            If Not IsEmpty(corrValue) Then
                If Not IsDuplicate(dataId1, dataId2, uniquePairs) Then
                    Set pairData = New Dictionary
                    With pairData
                        .Add "dataId1", dataId1
                        .Add "dataId2", dataId2
                        .Add "dataId", dataId1 & ":" & dataId2
                        .Add "corr", corrValue
                    End With
                    correlationData.Add pairData
                    uniquePairs.Add dataId1 & ":" & dataId2, Nothing
                End If
            End If
            DoEvents
        Next hCell
        DoEvents
    Next vCell

    CorrJsonv2 = JsonConverter.ConvertToJson(correlationData)
End Function


