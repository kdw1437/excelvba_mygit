Attribute VB_Name = "PostCorr"
Function GenerateJsonString() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Market Data")

    Dim DataString As String
    DataString = "[" ' �� json array�� DataString�� �ʱ�ȭ
    Dim uniquePairs As New Scripting.Dictionary
    Dim VerticalRange As Range, HorizontalRange As Range
    Set VerticalRange = ws.Range("M8", ws.Range("M8").End(xlDown))
    Set HorizontalRange = ws.Range("O7", ws.Range("O7").End(xlToRight))

    Dim vCell As Range, hCell As Range
    Dim jsonObject As String, comma As String
    comma = ""

    For Each vCell In VerticalRange
        For Each hCell In HorizontalRange
            Dim dataId1 As String
            Dim dataId2 As String
            Dim corrValue As Variant

            dataId1 = vCell.value
            dataId2 = hCell.value
            corrValue = ws.Cells(vCell.Row, hCell.Column).value

            If Not IsEmpty(corrValue) Then
                ' pair�� �ݺ����� Ȯ��
                If Not IsDuplicate(dataId1, dataId2, uniquePairs) Then
                    jsonObject = comma & "{""dataId1"": """ & dataId1 & _
                                """, ""dataId2"": """ & dataId2 & _
                                """, ""dataId"": """ & dataId1 & ":" & dataId2 & _
                                """, ""corr"": " & corrValue & "}"
                    DataString = DataString & jsonObject
                    comma = ", "
                    ' ��(pair)�� duplicate�� �����ϱ� ���� uniquePairs dictionary�� �߰�
                    uniquePairs(dataId1 & ":" & dataId2) = True
                End If
            End If
        Next hCell
    Next vCell

    DataString = DataString & "]" ' Close the JSON array
    GenerateJsonString = DataString
End Function

Sub PrintJsonString()
    Debug.Print GenerateJsonString()
End Sub


