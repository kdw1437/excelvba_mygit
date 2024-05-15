VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAsyncHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: CAsyncHandler
Public Sub HandleStateChange(xmlhttp As WinHttp.WinHttpRequest)
    If xmlhttp.readyState = 4 Then  ' Check if the request is complete
        If xmlhttp.Status = 200 Then
            MsgBox xmlhttp.responseText  ' If the request was successful, output the response
        Else
            MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.statusText  ' If the request failed, output the status
        End If
    End If
End Sub
