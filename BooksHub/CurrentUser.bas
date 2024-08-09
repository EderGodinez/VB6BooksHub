Attribute VB_Name = "CurrentUser"
Option Explicit
Private CurrentUser As User

Public Sub SetCurrentUser(book As User)
    Set CurrentUser = book
End Sub

Public Function GetCurrentUser() As User
    Set GetCurrentUser = CurrentUser
End Function


