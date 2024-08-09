VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form AdminForm 
   Caption         =   "Administrador"
   ClientHeight    =   10200
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   17685
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   17685
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdClear 
      Caption         =   "Limpiar"
      Height          =   615
      Left            =   7560
      TabIndex        =   25
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "Actualizar"
      Height          =   615
      Left            =   9840
      TabIndex        =   24
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox TxtId 
      Height          =   285
      Left            =   1680
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox CbCategory 
      DataSource      =   "DB"
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "Form1.frx":0000
      Left            =   7800
      List            =   "Form1.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   7
      Text            =   "-------------------------------------------------------------------"
      Top             =   3360
      Width           =   4455
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   480
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxtEditorial 
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   4080
      Width           =   4455
   End
   Begin VB.TextBox TxtPages 
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   6
      Top             =   2280
      Width           =   4455
   End
   Begin VB.TextBox TxtSearchBook 
      Height          =   375
      Left            =   12480
      TabIndex        =   14
      Top             =   4920
      Width           =   4455
   End
   Begin VB.TextBox TxtYear 
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox TxtISBN 
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3240
      Width           =   4455
   End
   Begin VB.TextBox TxtTitle 
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox TxtAuthor 
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2280
      Width           =   4455
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Eliminar"
      Height          =   615
      Left            =   14640
      TabIndex        =   9
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton CmdAddBock 
      Caption         =   "Agregar libro"
      Height          =   615
      Left            =   12120
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
   Begin MSComctlLib.ListView BookList 
      Height          =   4455
      Left            =   360
      TabIndex        =   0
      Top             =   5520
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   7858
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lB_page_error 
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   22
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label Lb_year_error 
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   21
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Image ImgPoster 
      BorderStyle     =   1  'Fixed Single
      Height          =   2775
      Left            =   13440
      MousePointer    =   4  'Icon
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label LbEditorial 
      Alignment       =   2  'Center
      Caption         =   "Editorial"
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   20
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label LbCategory 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Genero"
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6930
      TabIndex        =   19
      Top             =   3360
      Width           =   705
   End
   Begin VB.Label LbPages 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Paginas"
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6915
      TabIndex        =   18
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label LbYear 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Año"
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7095
      TabIndex        =   17
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label LbPortada 
      Alignment       =   2  'Center
      Caption         =   "Portada"
      BeginProperty Font 
         Name            =   "Sitka Display"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13920
      TabIndex        =   16
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label LbSearch 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Buscar libro:"
      BeginProperty Font 
         Name            =   "Sitka Display"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11160
      TabIndex        =   15
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label LbISBN 
      Alignment       =   2  'Center
      Caption         =   "ISBN"
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label LbAutor 
      Alignment       =   2  'Center
      Caption         =   "Autor"
      BeginProperty Font 
         Name            =   "Sitka Banner"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label LbTitle 
      Alignment       =   2  'Center
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "Sitka Display"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label LbTitlePage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bienvenido ADMIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   10
      Top             =   120
      Width           =   5565
      WordWrap        =   -1  'True
   End
   Begin VB.Menu logout 
      Caption         =   "Cerrar sesion"
   End
End
Attribute VB_Name = "AdminForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim allGenres As Collection
Dim imgPath As String
Dim searchTerm As String


Private Sub CmdAddBock_Click()
    If ValidateTextBoxes() Then
        MsgBox "Libro " & TxtTitle.Text & " guardado con exito :D.", vbInformation
        RegisterBook
        ResetTextBoxes
    End If
End Sub
Private Sub BookList_ItemClick(ByVal item As MSComctlLib.ListItem)
    TxtId.Text = item.Text
    TxtTitle.Text = item.SubItems(1)
    TxtAuthor.Text = item.SubItems(2)
    TxtISBN.Text = item.SubItems(3)
    TxtEditorial.Text = item.SubItems(4)
    TxtYear.Text = item.SubItems(5)
    TxtPages.Text = item.SubItems(6)
    CbCategory.Text = item.SubItems(7)
    ImgPoster.Picture = LoadPicture(item.SubItems(8))
    imgPath = item.SubItems(8)
End Sub

Private Sub CmdClear_Click()
ResetTextBoxes
End Sub

Private Sub CmdDelete_Click()
Dim selectedItem As ListItem
    Dim itemId As String
    Dim db As DBConection
    Dim connection As ADODB.connection
    Dim command As ADODB.command
    Dim query As String
    Dim Index As Integer
    ' Verificar si hay un ítem seleccionado
    If BookList.selectedItem Is Nothing Then
        MsgBox "Por favor, selecciona un libro para eliminar.", vbExclamation
        Exit Sub
    End If

    ' Obtener el ítem seleccionado
    Set selectedItem = BookList.selectedItem
    itemId = selectedItem.Text ' Suponiendo que el ID está en la primera columna

    ' Confirmar eliminación
    If MsgBox("¿Estás seguro de que deseas eliminar el libro con ID " & itemId & "?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    Index = selectedItem.Index
    If Index >= 1 And Index <= BookList.ListItems.Count Then
        BookList.ListItems.Remove Index
    Else
        MsgBox "Ítem no encontrado.", vbExclamation
        Exit Sub
    End If
    ' Configurar cadena de conexión y comando
    Set db = New DBConection
    Set connection = db.GetConnection
    Set command = New ADODB.command
    On Error GoTo ErrorHandler
    ' Configurar el comando de eliminación
    query = "DELETE FROM Books WHERE Id = ?"
    With command
        .ActiveConnection = connection
        .CommandText = query
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter(, adInteger, adParamInput, , itemId)
        .Execute
    End With
    ' Cerrar la conexión
    connection.Close
    Set connection = Nothing
    Set command = Nothing

    MsgBox "Libro eliminado correctamente.", vbInformation
    ResetTextBoxes
    Exit Sub

ErrorHandler:
    MsgBox "Error al eliminar el libro: " & Err.Description, vbExclamation
    If Not connection Is Nothing Then
        If connection.State = adStateOpen Then connection.Close
    End If
    Set connection = Nothing
    Set command = Nothing
End Sub

Private Sub CmdUpdate_Click()
  If TxtId.Text <> "" Then
    Dim miLibro As libro
    Set miLibro = New libro
    miLibro.Id = CInt(TxtId.Text)
    miLibro.Titulo = TxtTitle.Text
    miLibro.AnioPublicacion = CInt(TxtYear.Text)
    miLibro.Autor = TxtAuthor.Text
    miLibro.Editorial = TxtEditorial.Text
    miLibro.Genero = CbCategory.Text
    miLibro.ISBN = TxtISBN.Text
    miLibro.NumeroPaginas = CInt(TxtPages.Text)
    miLibro.Poster = imgPath
    Dim resultado As String
    resultado = UpdateBook(miLibro)
    MsgBox resultado
    ResetTextBoxes
    miLibro.ObtenerDescripcion
    ChangeValueInListView miLibro
    Else
        MsgBox "No hay ID.", vbExclamation
    End If

End Sub

Private Sub Form_Load()
 Dim item As ListItem
    With BookList
        .View = lvwReport
        .ColumnHeaders.Add , , "Id", 1000
        .ColumnHeaders.Add , , "Titulo", 2300
        .ColumnHeaders.Add , , "Autor", 2300
        .ColumnHeaders.Add , , "ISBN", 2300
        .ColumnHeaders.Add , , "Editorial", 2300
        .ColumnHeaders.Add , , "Año", 1200
        .ColumnHeaders.Add , , "Paginas", 1500
        .ColumnHeaders.Add , , "Genero", 1800
        .ColumnHeaders.Add , , "Ruta imagen", 1800
    End With
    Set allGenres = New Collection
    Dim RsGeneros As ADODB.recordset
    On Error GoTo ErrorHandler
    Set RsGeneros = LoadDBGenders
    If RsGeneros Is Nothing Then
        MsgBox "No se pudo obtener el Recordset de géneros.", vbExclamation
        Exit Sub
    End If
    With RsGeneros
        Do While Not .EOF
            Dim genreName As String
            genreName = .Fields("Name").Value
            allGenres.Add genreName, CStr(.Fields("Id").Value)
            .MoveNext
        Loop
    End With
    RsGeneros.Close
    Set RsGeneros = Nothing
    Dim RsBooks As ADODB.recordset
    Set RsBooks = LoadBooksDB
    BookList.ListItems.Clear
      ' Agregar los elementos al ListView
    With RsBooks
        Do While Not .EOF
            Set item = BookList.ListItems.Add(, , .Fields("Id").Value)
            item.SubItems(1) = .Fields("Titulo").Value
            item.SubItems(2) = .Fields("Autor").Value
            item.SubItems(3) = .Fields("ISBN").Value
            item.SubItems(4) = .Fields("Editorial").Value
            item.SubItems(5) = .Fields("AnioPublicacion").Value
            item.SubItems(6) = .Fields("Paginas").Value
            item.SubItems(7) = .Fields("Name").Value
            item.SubItems(8) = .Fields("Portada").Value
            .MoveNext
        Loop
    End With
    FillComboBox ""
    Exit Sub
ErrorHandler:
    MsgBox "Error al cargar los datos: " & Err.Description, vbExclamation
    If Not RsGeneros Is Nothing Then
        If RsGeneros.State = adStateOpen Then RsGeneros.Close
    End If
    Set RsGeneros = Nothing
    If Not BookList Is Nothing Then Set BookList = Nothing
End Sub

Private Sub Form_Resize()
    Const MIN_WIDTH As Integer = 17500
    Const MIN_HEIGHT As Integer = 11000
    Const MAX_WIDTH As Integer = 18500
    Const MAX_HEIGHT As Integer = 12000
    If Me.Width < MIN_WIDTH Then
        Me.Width = MIN_WIDTH
    ElseIf Me.Width > MAX_WIDTH Then
        Me.Width = MAX_WIDTH
    End If
    If Me.Height < MIN_HEIGHT Then
        Me.Height = MIN_HEIGHT
    ElseIf Me.Height > MAX_HEIGHT Then
        Me.Height = MAX_HEIGHT
    End If
End Sub


Private Sub imgPoster_Click()
On Error GoTo ErrHandler
    CommonDialog.CancelError = True
    CommonDialog.filter = "Imágenes|*.bmp;*.jpg;*.jpeg;*.gif;*.png"
    CommonDialog.ShowOpen
    ' Asignar la imagen seleccionada al control de imagen
    ImgPoster.Picture = LoadPicture(CommonDialog.FileName)
    ' Guardar la ruta del archivo en la variable
    imgPath = CommonDialog.FileName
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub Logout_Click()
SetCurrentUser New User
    AdminForm.Hide
    LoginForm.Show
End Sub

Private Sub TxtPages_KeyPress(KeyAscii As Integer)
    Const MAX_PAGES As Integer = 10000
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        If KeyAscii = 13 Then
        End If
        Dim currentValue As String
        currentValue = TxtPages.Text
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            currentValue = currentValue & Chr(KeyAscii)
        End If
        If Val(currentValue) >= MAX_PAGES Then
        lB_page_error.Caption = "Un libro no puede tener más de " & MAX_PAGES & " páginas"
            Beep
            KeyAscii = 0
        End If
    Else
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub TxtSearchBook_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then ' Enter key
       
        searchTerm = Trim(TxtSearchBook.Text)
        Dim RsBooks As ADODB.recordset
        If searchTerm = "" Then
            Set RsBooks = LoadBooksDB
        Else
            Set RsBooks = FilterBooks(searchTerm)
        End If
        
        If Not RsBooks Is Nothing Then
            PopulateBookList RsBooks
            RsBooks.Close
            Set RsBooks = Nothing
        End If
    End If
End Sub

Private Sub TxtYear_KeyPress(KeyAscii As Integer)
    Lb_year_error.Caption = ""
    Dim currentValue As String
    Dim maxYear As Integer
    maxYear = Year(Date) + 1
    If ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
        currentValue = TxtYear.Text
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            currentValue = currentValue & Chr(KeyAscii)
        End If
        If Val(currentValue) <= 0 Or Val(currentValue) >= maxYear Then
            Lb_year_error.Caption = "Año invalido!!"
            Beep
            KeyAscii = 0
        End If
    Else
        Beep
        KeyAscii = 0
    End If
End Sub
Private Function ValidateTextBoxes() As Boolean
  Const MAX_PAGES As Integer = 10000
    Dim isValid As Boolean
    Dim currentYear As Integer
    isValid = True
    ' Validar el TextBox TxtTitle
    If Trim(TxtTitle.Text) = "" Then
        MsgBox "El título no puede estar vacío.", vbExclamation
        isValid = False
         Exit Function
    End If
    
    ' Validar el TextBox TxtAuthor
    If Trim(TxtAuthor.Text) = "" Then
        MsgBox "El autor no puede estar vacío.", vbExclamation
        isValid = False
         Exit Function
    End If
    
    ' Validar el TextBox TxtISBN
    If Trim(TxtISBN.Text) = "" Then
        MsgBox "El ISBN no puede estar vacío.", vbExclamation
        isValid = False
         Exit Function
    End If
    
    ' Validar el TextBox TxtEditorial
    If Trim(TxtEditorial.Text) = "" Then
        MsgBox "El Editorial no puede estar vacío.", vbExclamation
        isValid = False
         Exit Function
    End If
    
    ' Validar el TextBox TxtPages
    If Not IsNumeric(TxtPages.Text) Or Val(TxtPages.Text) <= 0 Or Val(TxtPages.Text) > MAX_PAGES Then
        MsgBox "El número de páginas debe ser un número mayor que 0 y menor o igual a " & MAX_PAGES & ".", vbExclamation
        isValid = False
         Exit Function
    End If
    lB_page_error.Caption = ""
    
    ' Validar el TextBox TxtYear
    currentYear = Year(Date) ' Obtiene el año actual
    If Not IsNumeric(TxtYear.Text) Or Val(TxtYear.Text) <= 0 Or Val(TxtYear.Text) >= currentYear + 1 Then
        MsgBox "El año debe ser un número mayor que 0 y menor que " & (currentYear + 1) & ".", vbExclamation
        isValid = False
         Exit Function
    End If
    Lb_year_error.Caption = ""
    'Validar que una opcion del comboBox haya sido seleccionada'
    If CbCategory.ListIndex = -1 Then
    MsgBox "Debe seleccionar una genero para el libro " & TxtTitle.Text & " .", vbExclamation
     isValid = False
    Exit Function
    End If
    'Validar que se haya elegido una imagen para el libro'
    If ImgPoster.Picture = 0 Then
    MsgBox "Debe seleccionar una imagen poster para el libro " & TxtTitle.Text & " .", vbExclamation
     isValid = False
    Exit Function
    End If
    ' Devolver el resultado de la validación
    ValidateTextBoxes = isValid
End Function


Private Sub ResetTextBoxes()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Text = ""
        End If
        If TypeOf ctrl Is ComboBox Then
        ctrl.ListIndex = -1
        End If
        If TypeOf ctrl Is Image Then
         ctrl.Picture = Nothing
        End If
    Next ctrl
End Sub
Private Sub RegisterBook()
    Dim nuevoLibro As New libro
    nuevoLibro.Titulo = TxtTitle.Text
    nuevoLibro.AnioPublicacion = TxtYear.Text
    nuevoLibro.Autor = TxtAuthor.Text
    nuevoLibro.Editorial = TxtEditorial.Text
    nuevoLibro.Genero = CbCategory.Text
    nuevoLibro.ISBN = TxtISBN.Text
    nuevoLibro.NumeroPaginas = TxtPages.Text
    nuevoLibro.Poster = ImgPoster.Picture
    Dim db As DBConection
    Dim connection As ADODB.connection
    Dim command As ADODB.command
    Dim recordset As ADODB.recordset
    Dim query As String
    Dim newBookId As Long
    Set db = New DBConection
    Set connection = db.GetConnection
    Set command = New ADODB.command
    Set recordset = New ADODB.recordset
    On Error GoTo ErrHandler
    query = "INSERT INTO Books (Titulo, AnioPublicacion, Autor, Editorial, Genero, ISBN, NumeroPaginas, Portada) " & _
        "OUTPUT INSERTED.Id " & _
        "VALUES (?, ?, ?, ?, (SELECT Id FROM Genders WHERE Name=?), ?, ?, ?)"

With command
    .ActiveConnection = connection
    .CommandText = query
    .CommandType = adCmdText
    .Parameters.Append .CreateParameter("Titulo", adVarChar, adParamInput, 255, nuevoLibro.Titulo)
    .Parameters.Append .CreateParameter("AnioPublicacion", adInteger, adParamInput, , nuevoLibro.AnioPublicacion)
    .Parameters.Append .CreateParameter("Autor", adVarChar, adParamInput, 255, nuevoLibro.Autor)
    .Parameters.Append .CreateParameter("Editorial", adVarChar, adParamInput, 255, nuevoLibro.Editorial)
    .Parameters.Append .CreateParameter("Genero", adVarChar, adParamInput, 255, nuevoLibro.Genero)
    .Parameters.Append .CreateParameter("ISBN", adVarChar, adParamInput, 255, nuevoLibro.ISBN)
    .Parameters.Append .CreateParameter("NumeroPaginas", adInteger, adParamInput, , nuevoLibro.NumeroPaginas)
    .Parameters.Append .CreateParameter("Portada", adVarChar, adParamInput, 255, imgPath)
End With

' Ejecutar la inserción y obtener el ID del nuevo libro
Set recordset = command.Execute
If Not recordset.EOF Then
    newBookId = recordset.Fields("Id").Value
End If

If Not recordset.EOF Then
    Dim item As ListItem
            Set item = BookList.ListItems.Add(, , newBookId)
            item.SubItems(1) = nuevoLibro.Titulo
            item.SubItems(2) = nuevoLibro.Autor
            item.SubItems(3) = nuevoLibro.ISBN
            item.SubItems(4) = nuevoLibro.Editorial
            item.SubItems(5) = nuevoLibro.AnioPublicacion
            item.SubItems(6) = nuevoLibro.NumeroPaginas
            item.SubItems(7) = nuevoLibro.Genero
            item.SubItems(8) = imgPath
End If
recordset.Close

' Cerrar conexión
connection.Close

Exit Sub
Cleanup:
    If Not connection Is Nothing Then
        If connection.State = adStateOpen Then connection.Close
    End If
    Set connection = Nothing
    Set command = Nothing

    Exit Sub

ErrHandler:
    MsgBox "Error al registrar el libro: " & Err.Description, vbExclamation
    Resume Cleanup
End Sub
Private Sub FillComboBox(filter As String)
     ' Limpia el ComboBox
    CbCategory.Clear
     If allGenres Is Nothing Then
        MsgBox "La colección de géneros no está inicializada.", vbExclamation
        Exit Sub
    End If
    Dim genre As Variant
    For Each genre In allGenres
        If filter = "" Or InStr(1, genre, filter, vbTextCompare) > 0 Then
            CbCategory.AddItem genre
        End If
    Next
End Sub
Public Function FilterBooks(searchTerm As String) As ADODB.recordset
    Dim connection As ADODB.connection
    Dim db As DBConection
    Dim query As String
    Set db = New DBConection
    Set connection = db.GetConnection
    On Error GoTo ErrorHandler
    Set FilterBooks = New ADODB.recordset
    query = "SELECT B.Id, Titulo, Autor, ISBN, Editorial, AnioPublicacion, NumeroPaginas as Paginas, G.Name,Portada " & _
            "FROM Books B INNER JOIN Genders G ON G.Id = B.Genero " & _
            "WHERE Titulo LIKE '%" & searchTerm & "%' OR Autor LIKE '%" & searchTerm & "%' " & _
            "OR ISBN LIKE '%" & searchTerm & "%' OR Editorial LIKE '%" & searchTerm & "%' OR G.Name LIKE '%" & searchTerm & "%'"
    FilterBooks.Open query, connection, adOpenStatic, adLockReadOnly
    Exit Function

ErrorHandler:
    MsgBox "Error al filtrar los libros: " & Err.Description, vbExclamation
    If Not FilterBooks Is Nothing Then
        If FilterBooks.State = adStateOpen Then FilterBooks.Close
    End If
    If Not connection Is Nothing Then
        If connection.State = adStateOpen Then connection.Close
    End If
    Set FilterBooks = Nothing
    Set connection = Nothing
End Function
Public Sub PopulateBookList(RsBooks As ADODB.recordset)
    BookList.ListItems.Clear
    With RsBooks
        Do While Not .EOF
            Dim item As ListItem
            Set item = BookList.ListItems.Add(, , .Fields("Id").Value)
            item.SubItems(1) = .Fields("Titulo").Value
            item.SubItems(2) = .Fields("Autor").Value
            item.SubItems(3) = .Fields("ISBN").Value
            item.SubItems(4) = .Fields("Editorial").Value
            item.SubItems(5) = .Fields("AnioPublicacion").Value
            item.SubItems(6) = .Fields("Paginas").Value
            item.SubItems(7) = .Fields("Name").Value
            item.SubItems(8) = .Fields("Portada").Value
            .MoveNext
        Loop
    End With
End Sub
Public Sub ChangeValueInListView(book As libro)
    Dim i As Integer
    Dim item As ListItem
    For i = 1 To BookList.ListItems.Count
        Set item = BookList.ListItems(i)
        If item.Text = CStr(book.Id) Then
            item.SubItems(1) = book.Titulo
            item.SubItems(2) = book.Autor
            item.SubItems(3) = book.ISBN
            item.SubItems(4) = book.Editorial
            item.SubItems(5) = CStr(book.AnioPublicacion)
            item.SubItems(6) = CStr(book.NumeroPaginas)
            item.SubItems(7) = book.Genero
            item.SubItems(8) = book.Poster
            Exit For
        End If
    Next i
End Sub
