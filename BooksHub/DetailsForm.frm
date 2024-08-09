VERSION 5.00
Begin VB.Form DetailsForm 
   Caption         =   "Detalles de libro"
   ClientHeight    =   11670
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   11670
   ScaleWidth      =   8835
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton AddFavorite 
      Caption         =   "Agregar a favoritos"
      Height          =   615
      Left            =   4800
      TabIndex        =   8
      Top             =   10680
      Width           =   3135
   End
   Begin VB.CommandButton CmdMarkRead 
      Caption         =   "Marcar como leído"
      Height          =   615
      Left            =   1680
      TabIndex        =   7
      Top             =   10680
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Paginas:"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   4920
      TabIndex        =   15
      Top             =   9840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Año:"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   1560
      TabIndex        =   14
      Top             =   9840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Genero:"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1560
      TabIndex        =   13
      Top             =   9000
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "ISBN:"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1560
      TabIndex        =   12
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Editorial:"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1440
      TabIndex        =   11
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Autor:"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1560
      TabIndex        =   10
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Titulo:"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1560
      TabIndex        =   9
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label LbGender 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   9000
      Width           =   5055
   End
   Begin VB.Label LbISBN 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   8160
      Width           =   5055
   End
   Begin VB.Label LbPages 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   9840
      Width           =   1695
   End
   Begin VB.Label LbYear 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   9840
      Width           =   1575
   End
   Begin VB.Label LbEditorial 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   7440
      Width           =   5055
   End
   Begin VB.Label LbAutor 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   6720
      Width           =   5055
   End
   Begin VB.Label LbTitle 
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   6000
      Width           =   5055
   End
   Begin VB.Image ImgPoster 
      Height          =   5655
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6135
   End
   Begin VB.Menu Back 
      Caption         =   "Volver"
   End
End
Attribute VB_Name = "DetailsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddFavorite_Click()
    Dim conn As ADODB.connection
    Dim cmd As ADODB.command
    Dim query As String
    Dim db As DBConection
    Dim currentUserId As Integer
    Dim currentBookId As Integer
    ' Obtener la conexión a la base de datos
    Set db = New DBConection
    Set conn = db.GetConnection
    currentUserId = GetCurrentUser.Id
    currentBookId = GetCurrentBook.Id
    query = "INSERT INTO BooksLikes (UserId, BookId) VALUES (?, ?)"
    Set cmd = New ADODB.command
    With cmd
        .ActiveConnection = conn
        .CommandText = query
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("UserId", adInteger, adParamInput, , currentUserId)
        .Parameters.Append .CreateParameter("BookId", adInteger, adParamInput, , currentBookId)
    End With
    On Error GoTo ErrorHandler
    cmd.Execute
    MsgBox "El libro " & GetCurrentBook.Titulo & " ha sido agregado a tus favoritos.", vbInformation
    conn.Close
    Set conn = Nothing
    Set cmd = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error al agregar a favorito: " & Err.Description, vbCritical
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set conn = Nothing
    Set cmd = Nothing
End Sub

Private Sub Back_Click()
DetailsForm.Hide
UsersViewForm.Show
End Sub

Private Sub CmdMarkRead_Click()
    Dim conn As ADODB.connection
    Dim cmd As ADODB.command
    Dim query As String
    Dim db As DBConection
    Dim currentUserId As Integer
    Dim currentBookId As Integer
    ' Obtener la conexión a la base de datos
    Set db = New DBConection
    Set conn = db.GetConnection
    currentUserId = GetCurrentUser.Id
    currentBookId = GetCurrentBook.Id
    query = "INSERT INTO BooksRead (UserId, BookId) VALUES (?, ?)"
    Set cmd = New ADODB.command
    With cmd
        .ActiveConnection = conn
        .CommandText = query
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("UserId", adInteger, adParamInput, , currentUserId)
        .Parameters.Append .CreateParameter("BookId", adInteger, adParamInput, , currentBookId)
    End With
    On Error GoTo ErrorHandler
    cmd.Execute
    MsgBox "El libro " & GetCurrentBook.Titulo & " ha sido marcado como leído.", vbInformation
    conn.Close
    Set conn = Nothing
    Set cmd = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error al marcar el libro como leído: " & Err.Description, vbCritical
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set conn = Nothing
    Set cmd = Nothing
End Sub

Private Sub Form_Resize()
    Me.Width = 9000
    Me.Height = 12435
End Sub

Private Sub Form_Load()
ImgPoster.Picture = LoadPicture(GetCurrentBook.Poster)
LbTitle.Caption = GetCurrentBook.Titulo
LbAutor.Caption = GetCurrentBook.Autor
LbYear.Caption = GetCurrentBook.AnioPublicacion
LbPages.Caption = GetCurrentBook.NumeroPaginas
LbGender.Caption = GetCurrentBook.Genero
LbEditorial.Caption = GetCurrentBook.Editorial
LbISBN.Caption = GetCurrentBook.ISBN
End Sub

