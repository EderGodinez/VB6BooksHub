VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ButtonHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents cmdButton As CommandButton
Attribute cmdButton.VB_VarHelpID = -1

Private Sub cmdButton_Click()
    Dim libro As libro
    Dim frm As Form
    Set libro = GetBookById(cmdButton.Tag)
    SetCurrentBook libro
    For Each frm In Forms
        If frm.Name <> "DetailsForm" Then
            Unload frm
        End If
    Next frm
    DetailsForm.Show
End Sub
