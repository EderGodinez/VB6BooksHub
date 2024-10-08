USE [master]
GO
/****** Object:  Database [BooksHub]    Script Date: 09/08/2024 12:31:07 a. m. ******/
CREATE DATABASE [BooksHub]
 CONTAINMENT = NONE
 ON  PRIMARY 
( NAME = N'BooksHub', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\DATA\BooksHub.mdf' , SIZE = 8192KB , MAXSIZE = UNLIMITED, FILEGROWTH = 65536KB )
 LOG ON 
( NAME = N'BooksHub_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL16.MSSQLSERVER\MSSQL\DATA\BooksHub_log.ldf' , SIZE = 8192KB , MAXSIZE = 2048GB , FILEGROWTH = 65536KB )
 WITH CATALOG_COLLATION = DATABASE_DEFAULT, LEDGER = OFF
GO
ALTER DATABASE [BooksHub] SET COMPATIBILITY_LEVEL = 160
GO
IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [BooksHub].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO
ALTER DATABASE [BooksHub] SET ANSI_NULL_DEFAULT OFF 
GO
ALTER DATABASE [BooksHub] SET ANSI_NULLS OFF 
GO
ALTER DATABASE [BooksHub] SET ANSI_PADDING OFF 
GO
ALTER DATABASE [BooksHub] SET ANSI_WARNINGS OFF 
GO
ALTER DATABASE [BooksHub] SET ARITHABORT OFF 
GO
ALTER DATABASE [BooksHub] SET AUTO_CLOSE OFF 
GO
ALTER DATABASE [BooksHub] SET AUTO_SHRINK OFF 
GO
ALTER DATABASE [BooksHub] SET AUTO_UPDATE_STATISTICS ON 
GO
ALTER DATABASE [BooksHub] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO
ALTER DATABASE [BooksHub] SET CURSOR_DEFAULT  GLOBAL 
GO
ALTER DATABASE [BooksHub] SET CONCAT_NULL_YIELDS_NULL OFF 
GO
ALTER DATABASE [BooksHub] SET NUMERIC_ROUNDABORT OFF 
GO
ALTER DATABASE [BooksHub] SET QUOTED_IDENTIFIER OFF 
GO
ALTER DATABASE [BooksHub] SET RECURSIVE_TRIGGERS OFF 
GO
ALTER DATABASE [BooksHub] SET  DISABLE_BROKER 
GO
ALTER DATABASE [BooksHub] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO
ALTER DATABASE [BooksHub] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO
ALTER DATABASE [BooksHub] SET TRUSTWORTHY OFF 
GO
ALTER DATABASE [BooksHub] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO
ALTER DATABASE [BooksHub] SET PARAMETERIZATION SIMPLE 
GO
ALTER DATABASE [BooksHub] SET READ_COMMITTED_SNAPSHOT OFF 
GO
ALTER DATABASE [BooksHub] SET HONOR_BROKER_PRIORITY OFF 
GO
ALTER DATABASE [BooksHub] SET RECOVERY SIMPLE 
GO
ALTER DATABASE [BooksHub] SET  MULTI_USER 
GO
ALTER DATABASE [BooksHub] SET PAGE_VERIFY CHECKSUM  
GO
ALTER DATABASE [BooksHub] SET DB_CHAINING OFF 
GO
ALTER DATABASE [BooksHub] SET FILESTREAM( NON_TRANSACTED_ACCESS = OFF ) 
GO
ALTER DATABASE [BooksHub] SET TARGET_RECOVERY_TIME = 60 SECONDS 
GO
ALTER DATABASE [BooksHub] SET DELAYED_DURABILITY = DISABLED 
GO
ALTER DATABASE [BooksHub] SET ACCELERATED_DATABASE_RECOVERY = OFF  
GO
ALTER DATABASE [BooksHub] SET QUERY_STORE = ON
GO
ALTER DATABASE [BooksHub] SET QUERY_STORE (OPERATION_MODE = READ_WRITE, CLEANUP_POLICY = (STALE_QUERY_THRESHOLD_DAYS = 30), DATA_FLUSH_INTERVAL_SECONDS = 900, INTERVAL_LENGTH_MINUTES = 60, MAX_STORAGE_SIZE_MB = 1000, QUERY_CAPTURE_MODE = AUTO, SIZE_BASED_CLEANUP_MODE = AUTO, MAX_PLANS_PER_QUERY = 200, WAIT_STATS_CAPTURE_MODE = ON)
GO
USE [BooksHub]
GO
/****** Object:  Table [dbo].[Books]    Script Date: 09/08/2024 12:31:07 a. m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Books](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[Titulo] [nvarchar](255) NOT NULL,
	[AnioPublicacion] [int] NOT NULL,
	[Autor] [nvarchar](255) NOT NULL,
	[Editorial] [nvarchar](255) NOT NULL,
	[Genero] [int] NULL,
	[ISBN] [nvarchar](255) NOT NULL,
	[NumeroPaginas] [int] NOT NULL,
	[Portada] [nvarchar](255) NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[BooksLikes]    Script Date: 09/08/2024 12:31:08 a. m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[BooksLikes](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[UserId] [int] NULL,
	[BookId] [int] NULL,
PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[BooksRead]    Script Date: 09/08/2024 12:31:08 a. m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[BooksRead](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[UserId] [int] NULL,
	[BookId] [int] NULL,
PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Genders]    Script Date: 09/08/2024 12:31:08 a. m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Genders](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[Name] [varchar](20) NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Users]    Script Date: 09/08/2024 12:31:08 a. m. ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Users](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[Name] [varchar](50) NOT NULL,
	[Password] [varchar](150) NOT NULL,
	[Username] [varchar](50) NOT NULL,
	[Role] [varchar](10) NULL,
PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
SET IDENTITY_INSERT [dbo].[Books] ON 

INSERT [dbo].[Books] ([Id], [Titulo], [AnioPublicacion], [Autor], [Editorial], [Genero], [ISBN], [NumeroPaginas], [Portada]) VALUES (6, N'1984', 1949, N'George Orwell', N'Signet Classic', 5, N'978-0451524935', 328, N'C:\Users\Eder Godinez\Downloads\1984.jpg')
INSERT [dbo].[Books] ([Id], [Titulo], [AnioPublicacion], [Autor], [Editorial], [Genero], [ISBN], [NumeroPaginas], [Portada]) VALUES (8, N'El Gran Gatsby', 1925, N'F. Scott Fitzgerald', N'Scribner', 21, N'978-0743273565', 180, N'C:\Users\Eder Godinez\Downloads\el_gran_gatsby.jpg')
INSERT [dbo].[Books] ([Id], [Titulo], [AnioPublicacion], [Autor], [Editorial], [Genero], [ISBN], [NumeroPaginas], [Portada]) VALUES (9, N'Cien Años de Soledad', 1967, N'Gabriel García Márquez', N'Editorial Sudamericana', 3, N'978-0307474728', 417, N'C:\Users\Eder Godinez\Downloads\cien_anos_de_soledad.jpg')
INSERT [dbo].[Books] ([Id], [Titulo], [AnioPublicacion], [Autor], [Editorial], [Genero], [ISBN], [NumeroPaginas], [Portada]) VALUES (10, N'Matar a un Ruiseñor', 1960, N'Harper Lee', N'J. B. Lippincott & Co.', 3, N'978-0061120084', 281, N'C:\Users\Eder Godinez\Downloads\matar_a_un_ruisenor.jpg')
INSERT [dbo].[Books] ([Id], [Titulo], [AnioPublicacion], [Autor], [Editorial], [Genero], [ISBN], [NumeroPaginas], [Portada]) VALUES (11, N'Orgullo y Prejuicio', 1813, N'Jane Austen', N'T. Egerton', 6, N'978-0141439518', 279, N'C:\Users\Eder Godinez\Downloads\orgullo_y_prejuicio.jpg')
INSERT [dbo].[Books] ([Id], [Titulo], [AnioPublicacion], [Autor], [Editorial], [Genero], [ISBN], [NumeroPaginas], [Portada]) VALUES (12, N'El Señor de los Anillos', 1954, N'J. R. R. Tolkien', N'George Allen & Unwin', 8, N'978-0618640157', 1178, N'C:\Users\Eder Godinez\Downloads\el_senor_de_los_anillos.jpg')
INSERT [dbo].[Books] ([Id], [Titulo], [AnioPublicacion], [Autor], [Editorial], [Genero], [ISBN], [NumeroPaginas], [Portada]) VALUES (13, N'Crimen y Castigo', 1866, N'Fiódor Dostoyevski', N'The Russian Messenger', 18, N'978-0486415871', 671, N'C:\Users\Eder Godinez\Downloads\crimen_y_castigo.jpeg')
INSERT [dbo].[Books] ([Id], [Titulo], [AnioPublicacion], [Autor], [Editorial], [Genero], [ISBN], [NumeroPaginas], [Portada]) VALUES (14, N'El Principito', 1943, N'Antoine de Saint-Exupéry', N'Reynal & Hitchcock', 15, N'978-0156012195', 96, N'C:\Users\Eder Godinez\Downloads\el_principito.jpg')
INSERT [dbo].[Books] ([Id], [Titulo], [AnioPublicacion], [Autor], [Editorial], [Genero], [ISBN], [NumeroPaginas], [Portada]) VALUES (15, N'Drácula', 1897, N'Bram Stoker', N'Archibald Constable and Company', 19, N'978-0141439846', 418, N'C:\Users\Eder Godinez\Downloads\dracula.jpg')
INSERT [dbo].[Books] ([Id], [Titulo], [AnioPublicacion], [Autor], [Editorial], [Genero], [ISBN], [NumeroPaginas], [Portada]) VALUES (16, N'Don Quijote de la Mancha', 1605, N'Miguel de Cervantes', N'Francisco de Robles', 21, N'978-0060934347', 1072, N'C:\Users\Eder Godinez\Downloads\don_quijote.jpg')
SET IDENTITY_INSERT [dbo].[Books] OFF
GO
SET IDENTITY_INSERT [dbo].[BooksLikes] ON 

INSERT [dbo].[BooksLikes] ([Id], [UserId], [BookId]) VALUES (5, 1, 8)
INSERT [dbo].[BooksLikes] ([Id], [UserId], [BookId]) VALUES (1, 1, 9)
INSERT [dbo].[BooksLikes] ([Id], [UserId], [BookId]) VALUES (6, 3, 6)
INSERT [dbo].[BooksLikes] ([Id], [UserId], [BookId]) VALUES (2, 3, 10)
INSERT [dbo].[BooksLikes] ([Id], [UserId], [BookId]) VALUES (4, 3, 11)
INSERT [dbo].[BooksLikes] ([Id], [UserId], [BookId]) VALUES (3, 5, 13)
SET IDENTITY_INSERT [dbo].[BooksLikes] OFF
GO
SET IDENTITY_INSERT [dbo].[BooksRead] ON 

INSERT [dbo].[BooksRead] ([Id], [UserId], [BookId]) VALUES (1, 1, 6)
INSERT [dbo].[BooksRead] ([Id], [UserId], [BookId]) VALUES (4, 1, 14)
INSERT [dbo].[BooksRead] ([Id], [UserId], [BookId]) VALUES (6, 3, 6)
INSERT [dbo].[BooksRead] ([Id], [UserId], [BookId]) VALUES (5, 3, 9)
INSERT [dbo].[BooksRead] ([Id], [UserId], [BookId]) VALUES (2, 3, 12)
INSERT [dbo].[BooksRead] ([Id], [UserId], [BookId]) VALUES (3, 5, 15)
SET IDENTITY_INSERT [dbo].[BooksRead] OFF
GO
SET IDENTITY_INSERT [dbo].[Genders] ON 

INSERT [dbo].[Genders] ([Id], [Name]) VALUES (4, N'Acción')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (17, N'Autoayuda')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (7, N'Aventura')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (11, N'Biografía')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (5, N'Ciencia Ficción')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (21, N'Clásico')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (2, N'Comedia')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (26, N'Crimen')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (22, N'Distopía')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (3, N'Drama')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (14, N'Educativo')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (24, N'Erótico')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (23, N'Espiritualidad')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (8, N'Fantasía')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (28, N'Filosofía')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (10, N'Histórico')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (19, N'Horror')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (30, N'Humor')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (15, N'Infantil')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (16, N'Juvenil')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (9, N'Misterio')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (12, N'Poesía')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (18, N'Policíaco')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (6, N'Romance')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (20, N'Satírico')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (13, N'Suspenso')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (1, N'Terror')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (27, N'Thriller')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (29, N'Viajes')
INSERT [dbo].[Genders] ([Id], [Name]) VALUES (25, N'Western')
SET IDENTITY_INSERT [dbo].[Genders] OFF
GO
SET IDENTITY_INSERT [dbo].[Users] ON 

INSERT [dbo].[Users] ([Id], [Name], [Password], [Username], [Role]) VALUES (1, N'Eder Godinez', N'~‘˜av‰qs¡', N'EderGS', N'user')
INSERT [dbo].[Users] ([Id], [Name], [Password], [Username], [Role]) VALUES (3, N'Juan Martinez', N'~‘˜av‰', N'Test', N'user')
INSERT [dbo].[Users] ([Id], [Name], [Password], [Username], [Role]) VALUES (4, N'Admin', N'~‘˜av‰qs¡', N'Admin', N'admin')
INSERT [dbo].[Users] ([Id], [Name], [Password], [Username], [Role]) VALUES (5, N'Test2', N'~‘˜av‰', N'Test2', N'user')
SET IDENTITY_INSERT [dbo].[Users] OFF
GO
/****** Object:  Index [UQ_BooksLikes_UserBook]    Script Date: 09/08/2024 12:31:08 a. m. ******/
ALTER TABLE [dbo].[BooksLikes] ADD  CONSTRAINT [UQ_BooksLikes_UserBook] UNIQUE NONCLUSTERED 
(
	[UserId] ASC,
	[BookId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
/****** Object:  Index [UQ_BooksRead_UserBook]    Script Date: 09/08/2024 12:31:08 a. m. ******/
ALTER TABLE [dbo].[BooksRead] ADD  CONSTRAINT [UQ_BooksRead_UserBook] UNIQUE NONCLUSTERED 
(
	[UserId] ASC,
	[BookId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [UQ__Genders__737584F644511501]    Script Date: 09/08/2024 12:31:08 a. m. ******/
ALTER TABLE [dbo].[Genders] ADD UNIQUE NONCLUSTERED 
(
	[Name] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [UQ__Username]    Script Date: 09/08/2024 12:31:08 a. m. ******/
ALTER TABLE [dbo].[Users] ADD  CONSTRAINT [UQ__Username] UNIQUE NONCLUSTERED 
(
	[Username] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
ALTER TABLE [dbo].[Users] ADD  DEFAULT ('user') FOR [Role]
GO
ALTER TABLE [dbo].[Books]  WITH CHECK ADD FOREIGN KEY([Genero])
REFERENCES [dbo].[Genders] ([Id])
ON UPDATE CASCADE
ON DELETE SET NULL
GO
ALTER TABLE [dbo].[BooksLikes]  WITH CHECK ADD FOREIGN KEY([BookId])
REFERENCES [dbo].[Books] ([Id])
ON UPDATE CASCADE
ON DELETE SET NULL
GO
ALTER TABLE [dbo].[BooksLikes]  WITH CHECK ADD FOREIGN KEY([UserId])
REFERENCES [dbo].[Users] ([Id])
ON UPDATE CASCADE
ON DELETE SET NULL
GO
ALTER TABLE [dbo].[BooksRead]  WITH CHECK ADD FOREIGN KEY([BookId])
REFERENCES [dbo].[Books] ([Id])
ON UPDATE CASCADE
ON DELETE SET NULL
GO
ALTER TABLE [dbo].[BooksRead]  WITH CHECK ADD FOREIGN KEY([UserId])
REFERENCES [dbo].[Users] ([Id])
ON UPDATE CASCADE
ON DELETE SET NULL
GO
ALTER TABLE [dbo].[Users]  WITH CHECK ADD CHECK  (([Role]='admin' OR [Role]='user'))
GO
USE [master]
GO
ALTER DATABASE [BooksHub] SET  READ_WRITE 
GO
