# BookHub MEGA+ (Librería digital con Visual Basic 6) - (Sprint #5)

Desarrollado por: Eder Yair Godinez Salazar

# Descripcion del proyecto

Esta es una biblioteca digital donde se puede explorar una amplia gama de libros de diversos géneros literarios. En este nuevo sistema, es posible acceder al catálogo de libros, ya sea de manera individual o agrupada por género, y explorar todos los títulos disponibles. Además, cada libro cuenta con información completa, incluyendo autor, editorial entre otros detalles. Este sistema ofrece otras funciones interesantes, como la posibilidad de guardar libros en una lista personalizada de libros ya leidos por el usuario y ademas de guardarlos en un listadi de sus libros favoritos. También permite marcar libros que no sean de nuestro agrado,por parte del lado de los usuarios administradores cuentan con una sección dedicada para añadir nuevos títulos al sistema, así como para editar y eliminar los registros de la biblioteca.



# Objetivos

- Crear una base de datos utilizando SQL Server.
- Implementar una interfaz gráfica funcional en Windows mediante formularios en Visual Basic 6.
- Utilizar botones para la ejecución de funciones específicas.
- Enlazar la base de datos con el sistema utilizando clases.
- Guardar los datos generados en la base de datos de SQL Server.
- Crear perfiles de usuario para que cada persona pueda personalizar su experiencia dentro del sistema.

# Requerimientos Técnicos

- Visual Basic 6
- SQL Server Management Studio 20


# Imagenes de aplicacion
    
## Usuarios comunes
### Login

![image](https://github.com/EderGodinez/VB6BooksHub/blob/main/Screenshots/Login.png)

![image](https://github.com/EderGodinez/VB6BooksHub/blob/main/Screenshots/LoginUserCredentials.png)
### Visualizacion de libros

![image](https://github.com/EderGodinez/VB6BooksHub/blob/main/Screenshots/BookListToUsers.png)

### Detalles de libro
![image](https://github.com/EderGodinez/VB6BooksHub/blob/main/Screenshots/BookDetails.png)

### Libros favoritos 
![image](https://github.com/EderGodinez/VB6BooksHub/blob/main/Screenshots/FavoriteBooks.png)

### Libros leidos
![image](https://github.com/EderGodinez/VB6BooksHub/blob/main/Screenshots/ReadedBooks.png)

### Agregar un libro a leidos
![image](https://github.com/EderGodinez/VB6BooksHub/blob/main/Screenshots/ReadAction.png)

## Usuarios administradores
### Login

![image](https://github.com/EderGodinez/VB6BooksHub/blob/main/Screenshots/LoginAdminCredentials.png)

![image](https://github.com/EderGodinez/VB6BooksHub/blob/main/Screenshots/LoginMessageAccess.png)

### Tabla de gestion de libros

![image](https://github.com/EderGodinez/VB6BooksHub/blob/main/Screenshots/AdminForm.png)

![image](https://github.com/EderGodinez/VB6BooksHub/blob/main/Screenshots/GendersList.png)

### Busqueda de libros por medio de cualquier campo
![image](https://github.com/EderGodinez/VB6BooksHub/blob/main/Screenshots/SearchingBook.png)


# Instrucciones de Instalación

## Paso 1: Descargar y Configurar el Proyecto

1. **Clonar el Repositorio**:
   - Abre tu terminal o línea de comandos.
   - Posisionate en la carpeta en la que se guardara el proyecto
   - Ejecuta el siguiente comando para clonar el repositorio en tu máquina local:

     ```bash
     git clone https://github.com/EderGodinez/VB6BooksHub.git
     ```

2. **Instalar Visual Basic 6**:
   - Si no lo tienes instalado, descarga e instala Visual Basic 6.

3. **Abrir el Proyecto en Visual Basic 6**:
   - Abre Visual Basic 6.
   - Ve a la opción `Archivo` en la barra de menú.
   - Selecciona `Abrir archivo...`.
   - Navega hasta la ubicación donde descomprimiste el proyecto y abre el archivo `program1.vbp`.

¡Listo! El proyecto ahora está disponible en Visual Basic 6.

## Paso 2: Enlazar la Base de Datos

1. **Importar la Base de Datos**:
   - Importa los datos de la carpeta `BD` archivo ``BooksHub.sql`` en una base de datos de SQL Server.

2. **Configurar la Conexión en Visual Basic 6**:
   - En Visual Basic 6, dirígete a la sección de `Class Modules` .
   - Dentro del archivo `DBConection` Busca la línea de código que contiene la configuración de la conexión a la base de datos:

     ```vb
     Con.Open "Driver={SQL Server}; 
         Server= TU SERVIDOR; 
         Database= NOMBRE DE TU BASE DE DATOS; 
         User Id= TU USUARIO; 
         Password= TU CONTRASEÑA;"
     ```

   - Reemplaza `TU SERVIDOR`, `NOMBRE DE TU BASE DE DATOS`, `TU USUARIO`, y `TU CONTRASEÑA` con las credenciales correspondientes a tu configuración de SQL Server.

¡Listo! El proyecto estará conectado con tu base de datos SQL Server.

# Descripción de como se realizó

Para realizar esta fase del proyecto se realizaron los siguientes pasos: 
### Crear una base de datos en SQL Server y conectarla al sistema
Para mostrar los registros necesarios, y utilizar Visual Basic 6 para crear, administrar y administrar toda la funcionalidad necesaria. programa. Primero, configurar la base de datos fue fácil. El verdadero desafío fue conectar esta biblioteca a Visual Basic 6, que era un nuevo lenguaje para mí, pero con el que no estaba familiarizado. Con la ayuda de los cursos impartidos, fortalecí mis habilidades con bases de datos y adquirí los conocimientos necesarios para conectar bases de datos a sistemas a través de módulos.
Además, Visual Basic 6 introdujo un módulo para conectarse a una base de datos. Esto permite que cada formulario solo llame a una función del módulo para acceder a la base de datos. 
### Creacion de formularios
En este caso se utilizaron formularios y botones para empezar a diseñar una sesión de usuario personalizada. Esto es útil para desarrollar funciones futuras del sistema, ya que facilita el proceso de registro y visualización de datos según las preferencias del usuario. Luego el sistema implementa diversas funciones, como mostrar los libros disponibles y su información, marcarlos para su posterior lectura, agregarlos a favoritos.
# Diagrama Base de datos

![image](https://github.com/EderGodinez/VB6BooksHub/blob/main/DB/Dise%C3%B1oDB.png)

# Problemas conocidos
1. Escasa información de Visual Basic 6.
2. Dificultad con entender en algunos aspectos la sintaxis de visul basic 6.

# Restrospectiva
| Aspecto                    | Detalles                                                                                                                                                                                                                                                                                                               |
|--------------------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| ¿Qué salió bien?           | * Se pudiero agregar los componentes de visualizacion de manera dinamica en formulario de usuarios comunes <br>* Se pudo entender un poco las bases de lo que es el lenguaje visual basic 6. <br> * Se completo el objetivo al igual que implemento de manera exitosa un login con encriptacion de contraseñas |
| ¿Qué puedo hacer diferente? | * Dedicarle más tiempo a comprender Visual Basic. <br> * Reutilización de código para evitar muchos Form  |
| ¿Qué no salió bien?        |  * Falta de un diseño un poco mas amigable con el usuario.|




