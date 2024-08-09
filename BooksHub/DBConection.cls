VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBConection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private connectionString As String
Private conn As ADODB.connection

Public Sub ConnectDatabase()
    On Error GoTo ErrorHandler
    connectionString = "Provider=SQLOLEDB;Data Source=DESKTOP-S1I8691;Initial Catalog=BooksHub;Integrated Security=SSPI;"
    If conn Is Nothing Then
        Set conn = New ADODB.connection
    End If
    If conn.State = adStateClosed Then
        conn.connectionString = connectionString
        conn.Open
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error al conectar a la base de datos: " & Err.Description, vbExclamation
    Set conn = Nothing
End Sub

Public Sub Disconnect()
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then
            conn.Close
        End If
        Set conn = Nothing
    End If
End Sub

Public Function GetConnection() As ADODB.connection
    If conn Is Nothing Then
        Call ConnectDatabase
    End If
    Set GetConnection = conn
End Function

