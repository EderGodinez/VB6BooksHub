VERSION 5.00
Begin VB.Form LoginForm 
   Caption         =   "Iniciar sesion"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13005
   LinkTopic       =   "Form2"
   ScaleHeight     =   7230
   ScaleWidth      =   13005
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdLogin 
      Caption         =   "Iniciar sesion"
      Height          =   615
      Left            =   7440
      MaskColor       =   &H00FFC0C0&
      TabIndex        =   7
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Limpiar"
      Height          =   615
      Left            =   5040
      TabIndex        =   6
      Top             =   5160
      Width           =   1935
   End
   Begin VB.TextBox TxtPass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   4920
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3600
      Width           =   4935
   End
   Begin VB.TextBox TxtUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4920
      TabIndex        =   1
      Top             =   2400
      Width           =   4935
   End
   Begin VB.Label LbGoRegister 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Registrarse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   10920
      TabIndex        =   8
      Top             =   6720
      Width           =   1425
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   4440
      Width           =   4815
   End
   Begin VB.Label LbPass 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      TabIndex        =   4
      Top             =   3720
      Width           =   1485
   End
   Begin VB.Label LbUser 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   3
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Acceso a BookHub"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4920
      TabIndex        =   0
      Top             =   480
      Width           =   3105
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdClear_Click()
If (TxtUser(1).Text <> "" Or TxtPass(2).Text <> "") Then
    If MsgBox("¿Estas seguro de limpiar los campos?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    Else
    ResetControls
    End If
End If
End Sub

Private Sub CmdLogin_Click()
    If TxtUser(1).Text = "" Or TxtPass(2).Text = "" Then
        MsgBox "Campo username y contraseña son requeridos", vbExclamation, "Campos Vacíos"
    Exit Sub
    Else
    Dim User As User
    Set User = GetUserRole(TxtUser(1).Text, TxtPass(2).Text)
    SetCurrentUser User
    If User Is Nothing Then
        MsgBox "Nombre de usuario o contraseña incorrectos.", vbExclamation, "Error de inicio de sesión"
        Exit Sub
    End If
    Select Case User.Role
        Case "admin"
            MsgBox "Bienvenido a BooksHub, Administrador", vbInformation, "Inicio de sesión exitoso"
            ResetControls
            AdminForm.Show
        Case "user"
            MsgBox "Bienvenido a BooksHub", vbInformation, "Inicio de sesión exitoso"
           ResetControls
            UsersViewForm.Show
        Case Else
            MsgBox "Rol de usuario desconocido.", vbExclamation, "Error de inicio de sesión"
            ResetControls
            Exit Sub
    End Select
    LoginForm.Hide
    End If
End Sub
Private Function GetUserRole(Username As String, password As String) As User
 Dim query As String
    Dim cmd As ADODB.command
    Dim rs As ADODB.recordset
    Dim conn As ADODB.connection
    Dim db As DBConection
    Dim User As User
    On Error GoTo ErrorHandler
    
    ' Constante para la clave de encriptación
    Const ENCRYPTION_KEY As String = "M_e-AS:;hj+*bhs&5%?_!123!"
    
    ' Se encripta la contraseña para iniciar sesión
    password = Encrypt(password, ENCRYPTION_KEY)
    
    ' Inicializar los objetos
    Set db = New DBConection
    Set conn = db.GetConnection
    Set cmd = New ADODB.command
    Set rs = New ADODB.recordset

    ' Configurar el comando
    With cmd
        .ActiveConnection = conn
        .CommandText = "SELECT Id, Name, Username, Role FROM Users WHERE Username = ? AND Password = ?"
        .CommandType = adCmdText
        ' Agregar parámetros al comando
        .Parameters.Append .CreateParameter("Username", adVarChar, adParamInput, 255, Username)
        .Parameters.Append .CreateParameter("Password", adVarChar, adParamInput, 255, password)
    End With
    
    ' Ejecutar el comando y abrir el recordset
    rs.Open cmd
    
    ' Inicializar el objeto user
    Set User = Nothing
    
    ' Verificar si se encontró un registro
    If Not rs.EOF Then
        Set User = New User
        With User
            .Id = rs.Fields("Id").Value
            .Name = rs.Fields("Name").Value
            .Role = rs.Fields("Role").Value
            .Username = rs.Fields("Username").Value
        End With
    End If
    
    ' Cerrar el recordset y la conexión
    rs.Close
    conn.Close
    
    ' Limpiar objetos
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    Set db = Nothing
    
    ' Retornar el objeto user (o Nothing si no se encontró)
    Set GetUserRole = User
    
    Exit Function

ErrorHandler:
    MsgBox "Error al obtener el usuario: " & Err.Description, vbExclamation
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    Set db = Nothing
    Set GetUserRole = Nothing
End Function
Private Sub LbGoRegister_Click()
LoginForm.Hide
RegisterForm.Show
End Sub
Private Sub ResetControls()
TxtUser(1).Text = ""
TxtPass(2).Text = ""
End Sub
