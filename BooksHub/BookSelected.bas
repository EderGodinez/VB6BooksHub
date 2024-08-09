Attribute VB_Name = "BookSelected"
Option Explicit
Private CurrentBook As libro

Public Sub SetCurrentBook(book As libro)
    Set CurrentBook = book
End Sub

Public Function GetCurrentBook() As libro
    Set GetCurrentBook = CurrentBook
End Function


