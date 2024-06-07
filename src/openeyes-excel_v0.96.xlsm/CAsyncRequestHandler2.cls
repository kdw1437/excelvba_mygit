VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAsyncRequestHandler2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents xmlhttp As WinHttp.WinHttpRequest
Attribute xmlhttp.VB_VarHelpID = -1
Private isComplete As Boolean

' Method to send asynchronous POST request
Public Sub SendPostRequestAsync(DataString As String, url As String)
    Set xmlhttp = New WinHttp.WinHttpRequest
    isComplete = False
    With xmlhttp
        .Open "POST", url, True ' Asynchronous mode
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .Send "a=" & DataString
    End With
    ' Wait for the response
    Do While Not isComplete
        DoEvents ' Keep Excel responsive
    Loop
End Sub

' Event handler for state change
Private Sub xmlhttp_OnResponseFinished()
    If xmlhttp.Status = 200 Then
        MsgBox xmlhttp.responseText ' Show the server response
    Else
        MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.statusText
    End If
    isComplete = True
End Sub

' Event handler for ready state change
Private Sub xmlhttp_OnReadyStateChange()
    If xmlhttp.readyState = 4 Then ' Check if the request is complete
        If xmlhttp.Status = 200 Then
            MsgBox xmlhttp.responseText ' Show the server response
        Else
            MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.statusText
        End If
        isComplete = True
    End If
End Sub

