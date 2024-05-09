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
Private pVerticalRange As Range
Private pHorizontalRange As Range

' vertical range setter
Public Property Set VerticalRange(rng As Range)
    Set pVerticalRange = rng
End Property

' horizontal range setter
Public Property Set HorizontalRange(rng As Range)
    Set pHorizontalRange = rng
End Property

' Method to generate the JSON string
Public Function CorrJsonh() As String
    Dim DataString As String
    DataString = "[" ' Initialize the DataString as an empty JSON array
    Dim uniquePairs As New Scripting.Dictionary
    Dim vCell As Range, hCell As Range
    Dim jsonObject As String, comma As String
    comma = ""

    For Each vCell In pVerticalRange
        For Each hCell In pHorizontalRange
            Dim dataId1 As String
            Dim dataId2 As String
            Dim corrValue As Variant

            dataId1 = hCell.value
            dataId2 = vCell.value
            corrValue = vCell.Worksheet.Cells(vCell.row, hCell.Column).value

            If Not IsEmpty(corrValue) Then
                ' Check if the pair is a duplicate
                If Not IsDuplicate(dataId1, dataId2, uniquePairs) Then
                    jsonObject = comma & "{""dataId1"": """ & dataId1 & _
                                """, ""dataId2"": """ & dataId2 & _
                                """, ""dataId"": """ & dataId1 & ":" & dataId2 & _
                                """, ""corr"": " & corrValue & "}"
                    DataString = DataString & jsonObject
                    comma = ", "
                    ' Add the pair to the uniquePairs dictionary to track duplicates
                    uniquePairs(dataId1 & ":" & dataId2) = True
                End If
            End If
        Next hCell
    Next vCell

    DataString = DataString & "]" ' Close the JSON array
    CorrJsonh = DataString
End Function

' Method to generate the JSON string
Public Function CorrJsonv() As String
    Dim DataString As String
    DataString = "[" ' Initialize the DataString as an empty JSON array
    Dim uniquePairs As New Scripting.Dictionary
    Dim vCell As Range, hCell As Range
    Dim jsonObject As String, comma As String
    comma = ""

    For Each vCell In pVerticalRange
        For Each hCell In pHorizontalRange
            Dim dataId1 As String
            Dim dataId2 As String
            Dim corrValue As Variant

            dataId1 = vCell.value
            dataId2 = hCell.value
            corrValue = vCell.Worksheet.Cells(vCell.row, hCell.Column).value

            If Not IsEmpty(corrValue) Then
                ' Check if the pair is a duplicate
                If Not IsDuplicate(dataId1, dataId2, uniquePairs) Then
                    jsonObject = comma & "{""dataId1"": """ & dataId1 & _
                                """, ""dataId2"": """ & dataId2 & _
                                """, ""dataId"": """ & dataId1 & ":" & dataId2 & _
                                """, ""corr"": " & corrValue & "}"
                    DataString = DataString & jsonObject
                    comma = ", "
                    ' Add the pair to the uniquePairs dictionary to track duplicates
                    uniquePairs(dataId1 & ":" & dataId2) = True
                End If
            End If
        Next hCell
    Next vCell

    DataString = DataString & "]" ' Close the JSON array
    CorrJsonv = DataString
End Function


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
        Next hCell
    Next vCell

    CorrJsonh2 = JsonConverter.ConvertToJson(correlationData)
End Function

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
        Next hCell
    Next vCell

    CorrJsonv2 = JsonConverter.ConvertToJson(correlationData)
End Function


