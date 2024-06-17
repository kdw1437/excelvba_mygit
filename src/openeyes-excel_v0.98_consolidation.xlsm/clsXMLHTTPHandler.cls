VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsXMLHTTPHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public xmlhttp As MSXML2.XMLHTTP60

Public Sub SendPostRequest(DataString As String, url As String)
    Set xmlhttp = New MSXML2.XMLHTTP60
    xmlhttp.Open "POST", url, True  ' asynchronous를 위한 True
    xmlhttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xmlhttp.Send "a=" & DataString

    ' response check loop를 시작한다.
    CheckResponse
End Sub

' ready state를 확인하는 Method
Public Sub CheckResponse()
    On Error Resume Next
    Do While xmlhttp.readyState <> 4
        DoEvents  ' Excel의 응답성을 유지
    Loop
    On Error GoTo 0
    
    If xmlhttp.Status = 200 Then
        Debug.Print xmlhttp.ResponseText
        
        'MsgBox xmlhttp.responseText
        
        ' JSON response를 parsing 하고 Excel sheet에 값을 넣는 method 호출
        ParseAndWriteToSheet xmlhttp.ResponseText
    Else
        MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.StatusText
    End If
End Sub

' JSON response를 parsing 하고 Excel sheet에 값을 넣는 Method
Public Sub ParseAndWriteToSheet(JsonString As String)
    Dim json As Object
    Set json = JsonConverter.ParseJson(JsonString)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Quote")
    
    Dim i As Integer
    i = 10 ' B10부터 아래로 데이터를 입력

    Dim Item As Object
    For Each Item In json
        ws.Range("B" & i).value = Item("cntrCode")
        i = i + 1
    Next Item
End Sub
