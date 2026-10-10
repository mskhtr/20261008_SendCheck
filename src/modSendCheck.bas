Attribute VB_Name = "modSendCheck"
Option Explicit

Public Sub 送信前チェックを開く()

    On Error Resume Next
    Unload frmSendCheck
    On Error GoTo 0
    frmSendCheck.Show

End Sub
