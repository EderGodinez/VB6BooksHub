Attribute VB_Name = "DB_Actions"
Option Explicit
Dim db As DBConection
Public Function LoadBooksDB() As ADODB.recordset
    Dim RsBooks As ADODB.recordset
    Dim query As String
    On Error GoTo ErrorHandler
    Set db = New DBConection
    Dim conn As ADODB.connection
    Set conn = db.GetConnection
    Set RsBooks = New ADODB.recordset
    query = "SELECT B.Id, Titulo, Autor, ISBN, Editorial, AnioPublicacion, NumeroPaginas as Paginas, G.Name ,Portada " & _
            "FROM Books B INNER JOIN Genders G ON G.Id = B.Genero "
    RsBooks.Open query, conn, adOpenStatic, adLockReadOnly
    Set LoadBooksDB = RsBooks
    Exit Function
    ' Cerrar el recordset y la conexión
    RsBooks.Close
    conn.Close

    ' Limpiar los objetos
    Set RsBooks = Nothing
    Set conn = Nothing
    
    Exit Function

ErrorHandler:
    MsgBox "Error al cargar los libros: " & Err.Description, vbExclamation
    If Not RsBooks Is Nothing Then
        If RsBooks.State = adStateOpen Then RsBooks.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set RsBooks = Nothing
    Set conn = Nothing
End Function
Public Function PaginateBooksDB(Optional ByVal OffsetRows As Integer = 0) As ADODB.recordset
    Dim FetchRows As Integer
    FetchRows = 8
     Dim RsBooks As ADODB.recordset
    Dim query As String
    On Error GoTo ErrorHandler
    Set db = New DBConection
    Dim conn As ADODB.connection
    Set conn = db.GetConnection
    Set RsBooks = New ADODB.recordset
    ' Construir la consulta SQL
    query = "SELECT B.Id, Titulo, Autor, ISBN, Editorial, AnioPublicacion, NumeroPaginas as Paginas, G.Name ,Portada " & _
    " FROM Books B INNER JOIN Genders G " & _
    " ON G.Id = B.Genero " & _
    " ORDER BY B.Id OFFSET " & OffsetRows & " ROWS FETCH NEXT  8 ROWS ONLY"
    RsBooks.Open query, conn, adOpenStatic, adLockReadOnly
    Set PaginateBooksDB = RsBooks
    Exit Function
    RsBooks.Close
    conn.Close
    Set RsBooks = Nothing
    Set conn = Nothing
    
    Exit Function

ErrorHandler:
    MsgBox "Error al cargar los libros: " & Err.Description, vbExclamation
    If Not RsBooks Is Nothing Then
        If RsBooks.State = adStateOpen Then RsBooks.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set RsBooks = Nothing
    Set conn = Nothing
End Function

Public Function LoadDBGenders() As ADODB.recordset
    Dim query As String
    query = "SELECT * FROM Genders ORDER BY Id"
    On Error GoTo ErrorHandler
     If db Is Nothing Then
        Set db = New DBConection
    End If
    
    Dim conn As ADODB.connection
    Set conn = db.GetConnection
    Dim RsGeneros As ADODB.recordset
    Set RsGeneros = New ADODB.recordset
    RsGeneros.Open query, conn, adOpenStatic, adLockReadOnly
    Set LoadDBGenders = RsGeneros
    Exit Function
    RsGeneros.Close
    Set RsGeneros = Nothing
    conn.Close
    Set conn = Nothing
ErrorHandler:
    MsgBox "Error al cargar los géneros: " & Err.Description, vbExclamation
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set conn = Nothing
    Set RsGeneros = Nothing
End Function
Public Function UpdateBook(book As libro) As String
    Dim query As String
    Dim connection As ADODB.connection
    Dim command As ADODB.command
    Set db = New DBConection
    On Error GoTo ErrorHandler
    Set connection = db.GetConnection
    query = "UPDATE Books SET Titulo = ?, AnioPublicacion = ?, Autor = ?, Editorial = ?, " & _
            "Genero = (SELECT Id FROM Genders WHERE Name = ?), ISBN = ?, NumeroPaginas = ?, Portada = ? " & _
            "WHERE Id = ?"
    Set command = New ADODB.command
    With command
        .ActiveConnection = connection
        .CommandText = query
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("Titulo", adVarChar, adParamInput, 255, book.Titulo)
        .Parameters.Append .CreateParameter("AnioPublicacion", adInteger, adParamInput, , book.AnioPublicacion)
        .Parameters.Append .CreateParameter("Autor", adVarChar, adParamInput, 255, book.Autor)
        .Parameters.Append .CreateParameter("Editorial", adVarChar, adParamInput, 255, book.Editorial)
        .Parameters.Append .CreateParameter("Genero", adVarChar, adParamInput, 255, book.Genero)
        .Parameters.Append .CreateParameter("ISBN", adVarChar, adParamInput, 255, book.ISBN)
        .Parameters.Append .CreateParameter("NumeroPaginas", adInteger, adParamInput, , book.NumeroPaginas)
        .Parameters.Append .CreateParameter("Portada", adVarChar, adParamInput, 255, book.Poster)
        .Parameters.Append .CreateParameter("Id", adInteger, adParamInput, , book.Id)
    End With
    command.Execute
    connection.Close
    Set connection = Nothing
    Set command = Nothing
    UpdateBook = "Libro actualizado correctamente."
    Exit Function
ErrorHandler:
    UpdateBook = "Error al intentar actualizar libro con id: " & book.Id & " - " & Err.Description
    If Not connection Is Nothing Then
        If connection.State = adStateOpen Then connection.Close
    End If
    Set connection = Nothing
    Set command = Nothing
End Function
Public Function CountBooks() As Integer
    Dim query As String
    Dim connection As ADODB.connection
    Dim command As ADODB.command
    Dim rs As ADODB.recordset
    Dim TotalBooks As Integer
    Set db = New DBConection
    On Error GoTo ErrorHandler
    Set connection = db.GetConnection
    query = "SELECT COUNT(Id) as LibrosDisponibles FROM Books"
    Set command = New ADODB.command
    With command
        .ActiveConnection = connection
        .CommandText = query
    End With
    Set rs = command.Execute
    If Not rs.EOF Then
        TotalBooks = rs.Fields("LibrosDisponibles").Value
    End If
    connection.Close
    Set connection = Nothing
    Set command = Nothing
    CountBooks = TotalBooks
    Exit Function
ErrorHandler:
    MsgBox "Error al contar los libros: " & Err.Description, vbExclamation
    If Not connection Is Nothing Then
        If connection.State = adStateOpen Then connection.Close
    End If
    Set connection = Nothing
    Set command = Nothing
    Set rs = Nothing
    CountBooks = 0
End Function

Public Function GetBookById(Id As Integer) As libro
    Dim book As libro
    Set book = New libro
    Dim query As String
    Dim connection As ADODB.connection
    Dim command As ADODB.command
    Dim rs As ADODB.recordset
    Dim db As DBConection
    
    On Error GoTo ErrorHandler
    
    Set db = New DBConection
    Set connection = db.GetConnection
    
    query = "SELECT B.Id, Titulo, Autor, ISBN, Editorial, AnioPublicacion, NumeroPaginas AS Paginas, G.Name AS Genero, Portada " & _
            "FROM Books B INNER JOIN Genders G ON G.Id = B.Genero WHERE B.Id=" & Id

    Set command = New ADODB.command
    With command
        .ActiveConnection = connection
        .CommandText = query
    End With
    
    Set rs = command.Execute

    If rs Is Nothing Then
        MsgBox "El recordset no se inicializó."
        GoTo Cleanup
    End If
    
    If rs.EOF Then
        MsgBox "El libro con ID " & Id & " no se encontró."
        Set book = Nothing
    Else
        ' Establecer valores en el objeto libro
        With book
            .Poster = rs.Fields("Portada").Value
            .Titulo = rs.Fields("Titulo").Value
            .Id = rs.Fields("Id").Value
            .Autor = rs.Fields("Autor").Value
            .ISBN = rs.Fields("ISBN").Value
            .Editorial = rs.Fields("Editorial").Value
            .AnioPublicacion = rs.Fields("AnioPublicacion").Value
            .NumeroPaginas = rs.Fields("Paginas").Value
            .Genero = rs.Fields("Genero").Value
        End With
    End If

Cleanup:
    If Not connection Is Nothing Then
        If connection.State = adStateOpen Then connection.Close
    End If
    Set connection = Nothing
    Set command = Nothing
    Set rs = Nothing
    
    Set GetBookById = book
    Exit Function

ErrorHandler:
    MsgBox "Error al obtener el libro: " & Err.Description, vbExclamation
    Resume Cleanup
End Function
Public Function CountReadedBooks(UserId As Integer) As Integer
    Dim query As String
    Dim connection As ADODB.connection
    Dim command As ADODB.command
    Dim rs As ADODB.recordset
    Dim TotalBooks As Integer
    Set db = New DBConection
    On Error GoTo ErrorHandler
    Set connection = db.GetConnection
    query = "SELECT COUNT(Id) as LibrosLeidos FROM BooksRead WHERE UserId=" & UserId
    Set command = New ADODB.command
    With command
        .ActiveConnection = connection
        .CommandText = query
    End With
    Set rs = command.Execute
    If Not rs.EOF Then
        TotalBooks = rs.Fields("LibrosLeidos").Value
    End If
    connection.Close
    Set connection = Nothing
    Set command = Nothing
    CountReadedBooks = TotalBooks
    Exit Function
ErrorHandler:
    MsgBox "Error al contar los libros: " & Err.Description, vbExclamation
    If Not connection Is Nothing Then
        If connection.State = adStateOpen Then connection.Close
    End If
    Set connection = Nothing
    Set command = Nothing
    Set rs = Nothing
    CountReadedBooks = 0
End Function
Public Function CountLikeBooks(UserId As Integer) As Integer
    Dim query As String
    Dim connection As ADODB.connection
    Dim command As ADODB.command
    Dim rs As ADODB.recordset
    Dim TotalBooks As Integer
    Set db = New DBConection
    On Error GoTo ErrorHandler
    Set connection = db.GetConnection
    query = "SELECT COUNT(Id) as Favoritos FROM BooksLikes WHERE UserId=" & UserId
    Set command = New ADODB.command
    With command
        .ActiveConnection = connection
        .CommandText = query
    End With
    Set rs = command.Execute
    If Not rs.EOF Then
        TotalBooks = rs.Fields("Favoritos").Value
    End If
    connection.Close
    Set connection = Nothing
    Set command = Nothing
    CountLikeBooks = TotalBooks
    Exit Function
ErrorHandler:
    MsgBox "Error al contar los libros: " & Err.Description, vbExclamation
    If Not connection Is Nothing Then
        If connection.State = adStateOpen Then connection.Close
    End If
    Set connection = Nothing
    Set command = Nothing
    Set rs = Nothing
    CountLikeBooks = 0
End Function
Public Function GetLikedBooks(UserId As Integer) As ADODB.recordset
Dim RsBooks As ADODB.recordset
    Dim query As String
    On Error GoTo ErrorHandler
    Set db = New DBConection
    Dim conn As ADODB.connection
    Set conn = db.GetConnection
    Set RsBooks = New ADODB.recordset
    query = "SELECT B.Id, Titulo, Autor, ISBN, Editorial, AnioPublicacion, NumeroPaginas as Paginas, G.Name ,Portada " & _
            "FROM BooksLikes BL INNER JOIN Books B ON B.Id = BL.BookId " & _
            "INNER JOIN Genders G ON G.Id = B.Genero  " & _
            "WHERE BL.UserId=" & UserId
    RsBooks.Open query, conn, adOpenStatic, adLockReadOnly
    Set GetLikedBooks = RsBooks
    Exit Function
    ' Cerrar el recordset y la conexión
    RsBooks.Close
    conn.Close
    ' Limpiar los objetos
    Set RsBooks = Nothing
    Set conn = Nothing
    
    Exit Function

ErrorHandler:
    MsgBox "Error al cargar los libros: " & Err.Description, vbExclamation
    If Not RsBooks Is Nothing Then
        If RsBooks.State = adStateOpen Then RsBooks.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set RsBooks = Nothing
    Set conn = Nothing
End Function
Public Function GetReadedBooks(UserId As Integer) As ADODB.recordset
Dim RsBooks As ADODB.recordset
    Dim query As String
    On Error GoTo ErrorHandler
    Set db = New DBConection
    Dim conn As ADODB.connection
    Set conn = db.GetConnection
    Set RsBooks = New ADODB.recordset
    query = "SELECT B.Id, Titulo, Autor, ISBN, Editorial, AnioPublicacion, NumeroPaginas as Paginas, G.Name ,Portada " & _
            "FROM BooksRead BR INNER JOIN Books B ON B.Id = BR.BookId " & _
            "INNER JOIN Genders G ON G.Id = B.Genero  " & _
            "WHERE BR.UserId=" & UserId
    RsBooks.Open query, conn, adOpenStatic, adLockReadOnly
    Set GetReadedBooks = RsBooks
    Exit Function
    ' Cerrar el recordset y la conexión
    RsBooks.Close
    conn.Close

    ' Limpiar los objetos
    Set RsBooks = Nothing
    Set conn = Nothing
    
    Exit Function

ErrorHandler:
    MsgBox "Error al cargar los libros: " & Err.Description, vbExclamation
    If Not RsBooks Is Nothing Then
        If RsBooks.State = adStateOpen Then RsBooks.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set RsBooks = Nothing
    Set conn = Nothing
End Function


