VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAsyncRequestHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: CAsyncRequestHandler
Public WithEvents xmlhttp As WinHttp.WinHttpRequest
Attribute xmlhttp.VB_VarHelpID = -1

' Method to send asynchronous POST request
Public Sub SendPostRequestAsync(DataString As String, url As String)
    Set xmlhttp = New WinHttp.WinHttpRequest
    With xmlhttp
        .Open "POST", url, True ' Asynchronous mode
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .Send "a=" & DataString
    End With
End Sub

' Event handler for state change
Private Sub xmlhttp_OnResponseFinished()
    If xmlhttp.Status = 200 Then
        'MsgBox xmlhttp.responseText ' Show the server response
    Else
        MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.StatusText
    End If
End Sub

' Event handler for ready state change
Private Sub xmlhttp_OnReadyStateChange()
    If xmlhttp.readyState = 4 Then ' Check if the request is complete
        If xmlhttp.Status = 200 Then
            'MsgBox xmlhttp.responseText ' Show the server response
        Else
            MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.StatusText
        End If
    End If
End Sub
