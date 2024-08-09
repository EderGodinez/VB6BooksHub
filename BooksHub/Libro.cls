VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Libro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

    Public Id As Integer
    Public Titulo As String
    Public Autor As String
    Public ISBN As String
    Public Editorial As String
    Public AnioPublicacion As Integer
    Public NumeroPaginas As Integer
    Public Poster As String
    Public Genero As String
    
    

Private Sub Class_Initialize()
    ' Inicializar propiedades si es necesario
    Id = 0
    Titulo = ""
    Autor = ""
    ISBN = ""
    Editorial = ""
    AnioPublicacion = 0
    NumeroPaginas = 0
    Genero = ""
    Poster = ""
End Sub

' Método para establecer valores en las propiedades
Public Sub EstablecerValores(ByVal pPoster As String, ByVal pTitulo As String, ByVal pId As Integer, ByVal pAutor As String, ByVal pISBN As String, ByVal pEditorial As String, ByVal pAnioPublicacion As Integer, ByVal pNumeroPaginas As Integer, ByVal pGenero As String)
    Id = pId
    Titulo = pTitulo
    Autor = pAutor
    ISBN = pISBN
    Editorial = pEditorial
    AnioPublicacion = pAnioPublicacion
    NumeroPaginas = pNumeroPaginas
    Genero = pGenero
    Poster = pPoster
End Sub

' Método para obtener una descripción del libro
Public Function ObtenerDescripcion() As String
    ObtenerDescripcion = "Id: " & Id & " ,Título: " & Titulo & ", Autor: " & Autor & ", ISBN: " & ISBN & ", Editorial: " & Editorial & ", Año de Publicación: " & AnioPublicacion & ", Número de Páginas: " & NumeroPaginas & ", Género: " & Genero & " ,Portada" & Poster
End Function
