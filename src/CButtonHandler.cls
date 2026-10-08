VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Button As MSForms.CommandButton

Public ParentForm As Object

Public ActionName As String


Private Sub Button_Click()

    If Not ParentForm Is Nothing Then
        ParentForm.ButtonAction ActionName
    End If

End Sub
