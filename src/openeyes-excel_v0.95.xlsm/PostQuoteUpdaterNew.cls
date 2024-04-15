VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostQuoteUpdaterNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pWs As Worksheet
Private pRange As Range

'Getter for pWs
Public Property Get Worksheet() As Worksheet
    Set Worksheet = pWs
End Property

'Setter for pWs
Public Property Set Worksheet(value As Worksheet)
    Set pWs = value
End Property

'Getter for pRange
Public Property Get Range() As Range
    Set Range = pRange
End Property

'Setter for pRange
Public Property Set Range(value As Range)
    Set pRange = value
End Property

' Function to convert range to JSON string
Public Function ToJsonString() As String
    Dim dict As Object
    Dim cell As Range
    Dim row As Range
    Dim colHeaders As Range
    Dim jsonArr As Object
    Dim i As Long
    Dim c As Long
    Dim jsonStr As String
    
    ' Initialize JSON array
    Set jsonArr = CreateObject("Scripting.Dictionary")
    
    ' Set the range for column headers
    Set colHeaders = pWs.Range(pRange.Cells(1, 1), pRange.Cells(1, pRange.Columns.Count))
    
    ' Loop through each row in the range
    For Each row In pRange.Rows
        ' Skip the header row
        If row.row <> pRange.Rows(1).row Then
            ' Initialize a new dictionary for each row
            Set dict = CreateObject("Scripting.Dictionary")
            For c = 1 To pRange.Columns.Count
                ' Use column headers as keys and cell values as values
                ' If the cell is empty, add "null" as the value
                If Len(row.Cells(1, c).value) > 0 Then
                    dict(colHeaders.Cells(1, c).value) = row.Cells(1, c).value
                Else
                    dict(colHeaders.Cells(1, c).value) = Null
                End If
            Next c
            ' Add the dictionary to the JSON array
            jsonArr.Add "row" & i, dict
            i = i + 1
        End If
    Next row
    
    ' Convert JSON array to JSON string
    Dim Key As Variant
    jsonStr = "[" & vbCrLf
    For Each Key In jsonArr
        jsonStr = jsonStr & DictionaryToJSONString(jsonArr(Key)) & "," & vbCrLf ' Updated line
    Next Key
    ' Remove the last comma from the string
    If jsonArr.Count > 0 Then
        jsonStr = Left(jsonStr, Len(jsonStr) - 3) & vbCrLf ' Updated line
    End If
    jsonStr = jsonStr & "]"
    
    ToJsonString = jsonStr
End Function

' Helper function to convert a dictionary to JSON string
Function DictionaryToJSONString(dict As Object) As String
    Dim sc As Object
    Set sc = CreateObject("ScriptControl")
    sc.Language = "JScript"
    
    With sc
        .AddCode "function toJSON(obj) { return JSON.stringify(obj); }"
        DictionaryToJSONString = .Run("toJSON", dict)
    End With
    
    Set sc = Nothing
End Function
