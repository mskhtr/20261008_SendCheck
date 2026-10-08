VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCardClick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents CardFrame As MSForms.Frame
Public WithEvents CardLabel As MSForms.Label

Public ParentForm As Object


Private Sub CardFrame_Click()

    If ParentForm Is Nothing Then Exit Sub
    If CardFrame Is Nothing Then Exit Sub

    If Left(CardFrame.Name, 5) = "card_" Then
        ParentForm.ToggleCard CardFrame
    Else
        ParentForm.ToggleCategoryFilter CardFrame
    End If

End Sub


Private Sub CardLabel_Click()

    Dim card As Object

    If ParentForm Is Nothing Then Exit Sub
    If CardLabel Is Nothing Then Exit Sub

    If Right(CardLabel.Name, 6) = "_title" Or Right(CardLabel.Name, 6) = "_count" Then
        ParentForm.ToggleCategoryFilter CardLabel.Parent
        Exit Sub
    End If

    Set card = CardLabel.Parent
    If TypeName(card) <> "Frame" Then Exit Sub

    ParentForm.ToggleCard card

End Sub
