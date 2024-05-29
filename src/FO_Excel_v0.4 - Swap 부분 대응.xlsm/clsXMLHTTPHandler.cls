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
    xmlhttp.Open "POST", url, True  ' True for asynchronous
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xmlhttp.Send "a=" & DataString

    ' Start the response check loop
    CheckResponse
End Sub

' Method to poll the ready state
Public Sub CheckResponse()
    On Error Resume Next
    Do While xmlhttp.readyState <> 4
        DoEvents  ' Keep Excel responsive
    Loop
    On Error GoTo 0
    
    If xmlhttp.Status = 200 Then
        MsgBox xmlhttp.responseText
    Else
        MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.statusText
    End If
End Sub

