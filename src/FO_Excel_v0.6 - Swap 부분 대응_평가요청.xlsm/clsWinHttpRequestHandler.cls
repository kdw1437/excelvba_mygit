VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWinHttpRequestHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents WinHttp As WinHttp.WinHttpRequest
Attribute WinHttp.VB_VarHelpID = -1

Public Sub SendPostRequest(DataString As String, url As String)
    ' Create a new Win HTTP request
    Set WinHttp = New WinHttp.WinHttpRequest
    
    ' Set up the WinHttp object. Make an asynchronous POST request to the specified URL.
    WinHttp.Open "POST", url, True
    
    ' Set the request content-type header
    WinHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    ' Send the request with the DataString
    WinHttp.Send "a=" & DataString
End Sub

Private Sub WinHttp_OnResponseFinished()
    ' Check the status of the request
    If WinHttp.Status = 200 Then
        ' If the request was successful, output the response
        MsgBox WinHttp.responseText
    Else
        ' If the request failed, output the status
        MsgBox "Error: " & WinHttp.Status & " - " & WinHttp.statusText
    End If
End Sub

Private Sub WinHttp_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    MsgBox "Request error: " & ErrorNumber & " - " & ErrorDescription
End Sub

