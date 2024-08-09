VERSION 5.00
Begin VB.Form RegisterForm 
   Caption         =   "Registro"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   13125
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdRegister 
      Caption         =   "Registrar"
      Height          =   615
      Left            =   8040
      TabIndex        =   7
      Top             =   4440
      Width           =   3135
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Limpiar"
      Height          =   615
      Left            =   2400
      TabIndex        =   6
      Top             =   4440
      Width           =   3015
   End
   Begin VB.TextBox TxtCPass 
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
      IMEMode         =   3  'DISABLE
      Left            =   7560
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3120
      Width           =   4215
   End
   Begin VB.TextBox TxtPass 
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
      IMEMode         =   3  'DISABLE
      Left            =   7560
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox TxtUsername 
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
      Left            =   2160
      TabIndex        =   3
      Top             =   3120
      Width           =   3975
   End
   Begin VB.TextBox TxtName 
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
      Left            =   2160
      TabIndex        =   2
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label LbRegisterError 
      Height          =   495
      Left            =   4680
      TabIndex        =   12
      Top             =   5520
      Width           =   4095
   End
   Begin VB.Label LbConfrimPass 
      AutoSize        =   -1  'True
      Caption         =   "Confirmar contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7680
      TabIndex        =   11
      Top             =   2760
      Width           =   2205
   End
   Begin VB.Label LbPassword 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7680
      TabIndex        =   10
      Top             =   1200
      Width           =   1185
   End
   Begin VB.Label LbUserName 
      AutoSize        =   -1  'True
      Caption         =   "Usuario unico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2280
      TabIndex        =   9
      Top             =   2760
      Width           =   1440
   End
   Begin VB.Label LbName 
      AutoSize        =   -1  'True
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2280
      TabIndex        =   8
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label LbBackLogin 
      AutoSize        =   -1  'True
      Caption         =   "Iniciar sesion"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      MouseIcon       =   "RegisterForm.frx":0000
      TabIndex        =   1
      Top             =   6720
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Nuevo usuario"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   0
      Top             =   120
      Width           =   2115
   End
End
Attribute VB_Name = "RegisterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdClear_Click()
If (TxtPass.Text <> "" Or TxtCPass.Text <> "" Or TxtName.Text <> "" Or TxtUsername.Text <> "") Then
    If MsgBox("¿Estas seguro de limpiar los campos?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    Else
    ResetControls
    End If
End If
End Sub

Private Sub CmdRegister_Click()
    Dim username As String
    Dim password As String
    Dim confirmPassword As String
    Dim name As String
    Dim encryptedPassword As String
    Dim cmd As ADODB.command
    Dim db As DBConection
    Dim conn As ADODB.connection
    Const ENCRYPTION_KEY As String = "M_e-AS:;hj+*bhs&5%?_!123!"
    ' Obtener los valores del formulario
    username = TxtUsername.Text
    password = TxtPass.Text
    confirmPassword = TxtCPass.Text
    name = TxtName.Text
    ' Validar los datos ingresados
    If username = "" Or password = "" Or confirmPassword = "" Or name = "" Then
        MsgBox "Todos los campos deben ser completados.", vbExclamation, "Error de registro"
        Exit Sub
    End If
    'Validar si los campos de contraseñas son iguales
    
    If password <> confirmPassword Then
        MsgBox "Las contraseñas no coinciden.", vbExclamation, "Error de registro"
        Exit Sub
    End If
    encryptedPassword = Encrypt(password, ENCRYPTION_KEY)
    On Error GoTo ErrorHandler
    Set db = New DBConection
    Set conn = db.GetConnection

    Set cmd = New ADODB.command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "INSERT INTO Users (Username, Password, Name) VALUES (?, ?, ?)"
    cmd.Parameters.Append cmd.CreateParameter("Username", adVarChar, adParamInput, 255, username)
    cmd.Parameters.Append cmd.CreateParameter("Password", adVarChar, adParamInput, 255, encryptedPassword)
    cmd.Parameters.Append cmd.CreateParameter("Name", adVarChar, adParamInput, 255, name)
    cmd.Execute
    MsgBox "Usuario registrado correctamente.", vbInformation, "Registro exitoso"
    ResetControls
    conn.Close
    Set cmd = Nothing
    Set conn = Nothing
    RegisterForm.Hide
    LoginForm.Show
    Exit Sub
ErrorHandler:
    MsgBox "Error al registrar el usuario: " & Err.Description, vbExclamation
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set conn = Nothing
    Set cmd = Nothing
End Sub

Private Sub LbBackLogin_Click()
RegisterForm.Hide
LoginForm.Show
End Sub
Private Sub ResetControls()
    TxtUsername.Text = ""
    TxtPass.Text = ""
    TxtCPass.Text = ""
    TxtName.Text = ""
End Sub

