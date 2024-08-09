VERSION 5.00
Begin VB.Form UsersViewForm 
   Caption         =   "Coleccion de libros"
   ClientHeight    =   10005
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   14265
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   14265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdBack 
      Caption         =   "Anterior"
      Enabled         =   0   'False
      Height          =   495
      Left            =   11280
      TabIndex        =   1
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "Siguientes"
      Height          =   495
      Left            =   12720
      TabIndex        =   0
      Top             =   9360
      Width           =   1215
   End
   Begin VB.Menu Logout 
      Caption         =   "Cerrar sesion"
      Index           =   1
   End
   Begin VB.Menu BooksRead 
      Caption         =   "Leìdos"
   End
   Begin VB.Menu FavoriteBooks 
      Caption         =   "Favoritos"
   End
End
Attribute VB_Name = "UsersViewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Declare object variable as CommandButton and handle the events.'
Private buttonHandlers As Collection
Dim TotalBooks As Integer
Dim NumPage As Integer
Dim TotalPages As Integer

Private Sub BooksRead_Click()
Me.Hide
BooksReadForm.Show
End Sub

Private Sub CmdBack_Click()
NumPage = NumPage - 1
If NumPage = 0 Then
CmdBack.Enabled = False
End If
If ((NumPage + 1) < TotalPages) Then
    CmdNext.Enabled = True
End If
Dim booksArray() As Variant
Dim RsBooks As ADODB.recordset
Set RsBooks = PaginateBooksDB(NumPage * 8)
    If Not RsBooks Is Nothing Then
        If Not RsBooks.EOF And Not RsBooks.BOF Then
            ' Parsear los datos en un array
            booksArray = RsBooks.GetRows()
        Else
            MsgBox "No se encontraron libros en la base de datos.", vbExclamation, "Error"
        End If
    Else
        MsgBox "No se pudo cargar el recordset.", vbExclamation, "Error"
    End If
    RsBooks.Close
'Se resetan los controles para cambio de pagina
DestroyControls
CreateControls booksArray
End Sub

Private Sub CmdNext_Click()
NumPage = NumPage + 1
If ((NumPage + 1) = TotalPages) Then
    CmdNext.Enabled = False
End If
CmdBack.Enabled = True
Dim booksArray() As Variant
Dim RsBooks As ADODB.recordset
Set RsBooks = PaginateBooksDB(NumPage * 8)
    If Not RsBooks Is Nothing Then
        If Not RsBooks.EOF And Not RsBooks.BOF Then
            ' Parsear los datos en un array
            booksArray = RsBooks.GetRows()
        Else
            MsgBox "No se encontraron libros en la base de datos.", vbExclamation, "Error"
        End If
    Else
        MsgBox "No se pudo cargar el recordset.", vbExclamation, "Error"
    End If
    RsBooks.Close
'Se resetan los controles para cambio de pagina
DestroyControls
CreateControls booksArray
End Sub

Private Sub FavoriteBooks_Click()
Me.Hide
FavoriteBooksForm.Show
End Sub

Private Sub Form_Load()
'Obtener el numero de libros disponibles
TotalBooks = CountBooks()
    If TotalBooks < 8 Then
        CmdNext.Visible = False
        CmdBack.Visible = False
    End If
    'Se asignan el numero de paginas que existen
    TotalPages = TotalBooks \ 8
    'Si hay reciduo se aumenta ya que hay libros existentes
        If TotalBooks Mod 8 <> 0 Then
    TotalPages = TotalPages + 1
        End If
    'Obtener los libros
    NumPage = 0
    Dim booksArray() As Variant
    Dim RsBooks As ADODB.recordset
    Set RsBooks = PaginateBooksDB
    If Not RsBooks Is Nothing Then
        If Not RsBooks.EOF And Not RsBooks.BOF Then
            ' Parsear los datos en un array
            booksArray = RsBooks.GetRows()
        Else
            MsgBox "No se encontraron libros en la base de datos.", vbExclamation, "Error"
        End If
    Else
        MsgBox "No se pudo cargar el recordset.", vbExclamation, "Error"
    End If
    RsBooks.Close
    'Creacion de controls
    CreateControls booksArray
End Sub
Private Sub Logout_Click(Index As Integer)
SetCurrentUser New User
UsersViewForm.Hide
LoginForm.Show
End Sub
Private Sub CreateControls(booksArray() As Variant)
    Dim i As Integer
    Dim img As Image
    Dim lblAutor As Label
    Dim LbGender As Label
    Dim Id As TextBox
    'botenes
    Dim cmdObject As CommandButton
    Dim handler As ButtonHandler
    ' Crear una nueva colección para los manejadores de botones
    Set buttonHandlers = New Collection
    'Crear primera fila de cards de libros
    For i = 0 To UBound(booksArray, 2)
    'Salir en caso de que sean mas de 4 libros
     If i > 3 Then Exit For
        Set Id = Me.Controls.Add("VB.TextBox", "Id" & i)
        With Id
        .Text = booksArray(0, i)
        .Visible = False
        .Enabled = False
        End With
        Set img = Me.Controls.Add("VB.Image", "Img" & i)
        With img
            .Stretch = True
            .Left = ((i) * 3500) + 500
            .Top = 500
            .Width = 2800
            .Height = 3200
            .Picture = LoadPicture(booksArray(8, i)) ' Reemplaza con la ruta correcta
            .Visible = True
        End With
        ' Agregar primer label
        Set lblAutor = Me.Controls.Add("VB.Label", "LbAutor" & i)
        With lblAutor
            .Left = ((i) * 3500) + 500
            .Top = img.Top + img.Height + 100
            .Caption = "Autor: " & booksArray(2, i)
            .Visible = True
            .Width = img.Width
        End With
        ' Agregar segundo label
        Set LbGender = Me.Controls.Add("VB.Label", "LbGender" & i)
        With LbGender
            .Left = ((i) * 3500) + 500
            .Top = lblAutor.Top + lblAutor.Height
            .Caption = "Genero: " & booksArray(7, i)
            .Visible = True
        End With
        Set cmdObject = Me.Controls.Add("VB.CommandButton", "cmd" & i)
        With cmdObject
        .Left = ((i) * 3500) + img.Width - (cmdObject.Width / 1.75)
        .Top = LbGender.Top - 200
        .Caption = "Ver"
        .Visible = True
        .Tag = booksArray(0, i)
        End With
         Set handler = New ButtonHandler
        Set handler.cmdButton = cmdObject
        
        ' Almacenar el manejador en la colección
        buttonHandlers.Add handler, "Handler" & i
    Next i
    ' Crear la segunda fila de controles
   If UBound(booksArray, 2) > 4 Then
    For i = 4 To 7
        Set Id = Me.Controls.Add("VB.TextBox", "Id" & i)
        With Id
        .Text = i
        .Visible = False
        .Enabled = False
        End With
        Set img = Me.Controls.Add("VB.Image", "Img" & i)
        With img
            .Stretch = True
            .Left = ((i - 4) * 3500) + 500
            .Top = 5000
            .Width = 2800
            .Height = 3200
            .Picture = LoadPicture(booksArray(8, i)) ' Reemplaza con la ruta correcta
            .Visible = True
        End With
        Set lblAutor = Me.Controls.Add("VB.Label", "LbAutor" & i)
        With lblAutor
            .Left = ((i - 4) * 3500) + 500
            .Top = img.Top + img.Height + 100
            .Caption = "Autor: " & booksArray(2, i)
            .Visible = True
        End With
        ' Agregar segundo label
        Set LbGender = Me.Controls.Add("VB.Label", "LbGender" & i)
        With LbGender
            .Left = ((i - 4) * 3500) + 500
            .Top = lblAutor.Top + lblAutor.Height
            .Caption = "Genero: " & booksArray(7, i)
            .Visible = True
        End With
        ' Agregar boton
        Set cmdObject = Me.Controls.Add("VB.CommandButton", "cmd" & i)
        With cmdObject
        .Left = ((i - 4) * 3500) + img.Width - (cmdObject.Width / 1.75)
        .Top = LbGender.Top - 200
        .Caption = "Ver"
        .Visible = True
        .Tag = booksArray(0, i)
        End With
        Set handler = New ButtonHandler
        Set handler.cmdButton = cmdObject
        
        ' Almacenar el manejador en la colección
        buttonHandlers.Add handler, "Handler" & i
            Next i
      End If
End Sub
Private Sub DestroyControls()
    Dim ctrl As Control
    Dim i As Integer

    ' Recorrer los controles en el formulario
    For i = Me.Controls.Count - 1 To 0 Step -1
        Set ctrl = Me.Controls(i)

        ' Verificar si el control es uno de los creados dinámicamente
        If ctrl.Name Like "Id*" Or _
           ctrl.Name Like "Img*" Or _
           ctrl.Name Like "LbAutor*" Or _
           ctrl.Name Like "LbGender*" Or _
           ctrl.Name Like "cmd*" Then
            Me.Controls.Remove ctrl
        End If
    Next i

    ' Limpiar la colección de manejadores de botones
    If Not buttonHandlers Is Nothing Then
        Set buttonHandlers = Nothing
    End If
End Sub




