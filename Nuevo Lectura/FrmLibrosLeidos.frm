VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLibrosLeidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libros leidos"
   ClientHeight    =   6645
   ClientLeft      =   840
   ClientTop       =   855
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10695
   Begin VB.TextBox txtFecha_del_final_de_la_lectura 
      DataField       =   "Fecha_del_final_de_la_lectura"
      DataMember      =   "Libros"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2400
      TabIndex        =   34
      Top             =   5400
      Width           =   4575
   End
   Begin VB.TextBox txtFecha_de_inicio_de_la_lectura 
      DataField       =   "Fecha_de_inicio_de_la_lectura"
      DataMember      =   "Libros"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   2400
      TabIndex        =   32
      Top             =   4800
      Width           =   4575
   End
   Begin VB.TextBox txtSegundos_de_lectura 
      DataField       =   "Segundos_de_lectura"
      DataMember      =   "Libros"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   8760
      TabIndex        =   30
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtMinutos_de_lectura 
      DataField       =   "Minutos_de_lectura"
      DataMember      =   "Libros"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   8760
      TabIndex        =   28
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtHoras_de_lectura 
      DataField       =   "Horas_de_lectura"
      DataMember      =   "Libros"
      DataSource      =   "DataEnvironment1"
      Height          =   285
      Left            =   8760
      TabIndex        =   26
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton CMDPrimero 
      Height          =   495
      Left            =   3330
      Picture         =   "FrmLibrosLeidos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Ir al primer registro"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton CMDAnterior 
      Height          =   495
      Left            =   3810
      Picture         =   "FrmLibrosLeidos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Ir al registro anterior"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton CMDSiguiente 
      Height          =   495
      Left            =   5730
      Picture         =   "FrmLibrosLeidos.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Ir al siguiente registro"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton CmdUltimo 
      Height          =   495
      Left            =   6210
      Picture         =   "FrmLibrosLeidos.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Ir al ultimo registro (El más nuevo)"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton CMDBuscar1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      Picture         =   "FrmLibrosLeidos.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Buscar registros"
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton CMDNuevo 
      Height          =   495
      Left            =   840
      MouseIcon       =   "FrmLibrosLeidos.frx":0F32
      MousePointer    =   99  'Custom
      Picture         =   "FrmLibrosLeidos.frx":1084
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "0"
      ToolTipText     =   "Nuevo registro"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton CMDGuardar 
      Height          =   495
      Left            =   1560
      Picture         =   "FrmLibrosLeidos.frx":194E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Guardar registro"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton CMDEliminar 
      Height          =   495
      Left            =   7065
      MouseIcon       =   "FrmLibrosLeidos.frx":2218
      MousePointer    =   99  'Custom
      Picture         =   "FrmLibrosLeidos.frx":2522
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Eliminar registro"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton CMDEnd 
      Height          =   495
      Left            =   9360
      Picture         =   "FrmLibrosLeidos.frx":2DEC
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cerrar este formulario"
      Top             =   6000
      Width           =   495
   End
   Begin VB.TextBox txtAutor 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Autor_o_autores"
      DataMember      =   "Libros"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   0
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Text            =   "Una tabla aparte para el autor"
      Top             =   1080
      Width           =   8535
   End
   Begin VB.TextBox txtPaginasLen 
      BackColor       =   &H00C0C0C0&
      DataField       =   "CantidadDePaginas"
      DataMember      =   "Libros"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   0
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Text            =   "txtPaginasLen"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtAñoDePublicasion 
      BackColor       =   &H00C0C0C0&
      DataField       =   "AñoDePublicasion"
      DataMember      =   "Libros"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   1800
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Text            =   "Tabla aparte para el año"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox txtDireccionURL 
      BackColor       =   &H00C0C0C0&
      DataField       =   "DireccionURL"
      DataMember      =   "Libros"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   0
      MousePointer    =   99  'Custom
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Text            =   "txtDireccionURL"
      Top             =   2760
      Width           =   8535
   End
   Begin VB.TextBox txtTemaDelLibro 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Tema"
      DataMember      =   "Libros"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   3720
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Text            =   "Una tabla aparte para los temas"
      Top             =   1920
      Width           =   4815
   End
   Begin VB.TextBox txtNotas 
      BackColor       =   &H00C0C0C0&
      DataField       =   "Notas"
      DataMember      =   "Libros"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   975
      Left            =   0
      MousePointer    =   99  'Custom
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "FrmLibrosLeidos.frx":36B6
      Top             =   3600
      Width           =   8535
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5640
      Top             =   2400
   End
   Begin MSComctlLib.ProgressBar pgbMemo 
      Height          =   255
      Left            =   6600
      TabIndex        =   16
      Top             =   3360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   62000
   End
   Begin VB.ComboBox cmbNombreDelLibro 
      BackColor       =   &H00C0C0C0&
      DataField       =   "NombreDelLibro"
      DataMember      =   "Libros"
      DataSource      =   "DataEnvironment1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   360
      Left            =   0
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   360
      Width           =   8535
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fecha_del_final_de_la_lectura:"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   33
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Fecha_de_inicio_de_la_lectura:"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   31
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Segundos_de_lectura:"
      Height          =   255
      Index           =   2
      Left            =   8760
      TabIndex        =   29
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Minutos_de_lectura:"
      Height          =   255
      Index           =   1
      Left            =   8760
      TabIndex        =   27
      Top             =   1080
      Width           =   1440
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Horas_de_lectura:"
      Height          =   255
      Index           =   0
      Left            =   8760
      TabIndex        =   25
      Top             =   375
      Width           =   1335
   End
   Begin VB.Label lbId 
      BackColor       =   &H00808080&
      Caption         =   "lbId"
      DataField       =   "IdLibro"
      DataMember      =   "Libros"
      DataSource      =   "DataEnvironment1"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4305
      TabIndex        =   24
      Tag             =   "Id"
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label lbNombre 
      Caption         =   "Nombre del libro"
      Height          =   255
      Left            =   0
      MousePointer    =   99  'Custom
      TabIndex        =   23
      ToolTipText     =   "Copiar del porta papeles el nombre del libro."
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label lbAutor 
      Caption         =   "Nombre del autor o de los autores"
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label lbPaginasLen 
      BackColor       =   &H80000000&
      Caption         =   "Cantidad de paginas"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lbAñoDePublicasion 
      Caption         =   "Año de publicasion"
      Height          =   255
      Left            =   1800
      TabIndex        =   20
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lbTemaDelLibro 
      Caption         =   "Tema del libro"
      Height          =   255
      Left            =   3720
      TabIndex        =   19
      ToolTipText     =   "Ver lista de temas para elegir un tema."
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lbDireccionURL 
      Caption         =   "Direccion URL"
      Height          =   255
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   18
      ToolTipText     =   "Sitio donde se guarda el archivo en el disco duro."
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lbNotas 
      Caption         =   "Notas"
      Height          =   255
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   17
      Top             =   3360
      Width           =   2895
   End
End
Attribute VB_Name = "FrmLibrosLeidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum Colores
    Verde = &HFF00&
    Defauld = &H8000000F
End Enum

Public Function fNombreDelArchivo(Direccion_de_la_carpeta As Variant) As Variant  'Nombre del video que se esta reproducioendo. Es de solo lectura.
    'On Error GoTo AccionesCorrectivas
    Const cntSlash = "\"
    
    If Direccion_de_la_carpeta = "" Then
        Exit Function
    End If
    
    Dim vNombre As Variant, mNombre() As String
    
    mNombre = Split(Direccion_de_la_carpeta, cntSlash) 'Se obtiene el nombre con su extencion de archivo.
    vNombre = mNombre(UBound(mNombre)) 'Se carga en la variable.
    
    mNombre = Split(vNombre, ".") 'Despues se obtien el nombre sin su extencion de archivo.
    vNombre = mNombre(LBound(mNombre)) 'Se carga la nueva informacion en la variable.
    fNombreDelArchivo = vNombre 'Se retorna el nombre sin su extencion.
    
    'Exit Function
    'AccionesCorrectivas:
    'MsgBox Err.Description
    'MsgBox "Tengo problemas con prNombreDelVideo"
End Function


Public Sub sArrastrarUnArchivo(Data As Object, TextBox_para_la_direccion As Object, Control_para_el_nombre_del_archivo As Object) 'As Variant 'Para recuperar la direccion del archivo.
    On Error GoTo AccionesCorrectivas
    TextBox_para_la_direccion.Text = Data.Files.Item(1)
    Control_para_el_nombre_del_archivo.Text = fNombreDelArchivo(Data.Files.Item(1))
    Exit Sub
AccionesCorrectivas:
    'MsgBox "Tengo problemas con sArrastrarUnArchivo"
End Sub

Private Sub cAños_Actualizando()
'On Error Resume Next
'    txtAñoDePublicasion.Text = cAños.fGetF("AñoDePublicasion")
End Sub

Private Sub cAños_NuevoRegistro()
    txtAñoDePublicasion.Text = ""
End Sub

Private Sub cAños_SentFields()
'El guardado se debe realizar por aparte.

'On Error Resume Next
'    Dim SiNo As Integer
'    Dim cAñosLocal As clsDbx
'    Set cAñosLocal = New clsDbx
'    cAñosLocal.LoadDBConSQL ("Slect AñoDePublicasion from Años where AñoDePublicasion=" & txtAñoDePublicasion.Text) '("Select a.AñoDePublicasion from Años as a, Relacion r, Libros as L where L.IdLibro=" & cLibros.fGetF("IdLibro") & " And L.IdLibro=r.IdLibro And r.IdAño=a.IdAño order by a.IdAño")
'    cAñosLocal.mPrimero
'
'    SiNo = MsgBox("¿Desea guardar la lectura en progreso?", vbQuestion + vbYesNo)
'    If SiNo = vbYes Then
'        If cAñosLocal.fGetF("AñoDePublicasion") = "" Then
'            cAños.LetF txtAñoDePublicasion.Text, "AñoDePublicasion"
'        Else
'            MsgBox ("No se pudo guardar la informacion " & cAñosLocal.fGetF("AñoDePublicasion"))
'        End If
'    End If
End Sub

Private Sub cAutores_Actualizando()
'    txtAutor.Text = cAutores.fGetF("Autor_o_autores")
End Sub

Private Sub cAutores_NuevoRegistro()
    txtAutor.Text = ""
End Sub

Private Sub cAutores_SentFields()
'On Error Resume Next
'    Dim SiNo As Integer
'    Dim cAutoresLocal As clsDbx
'    Set cAutoresLocal = New clsDbx
'    cAutoresLocal.LoadDBConSQL ("select Autor_o_autores from Autores where Autor_o_autores=" & txtAutor.Text) '("Select Autors.Autor_o_autores from Autores as Autors, Relacion r, Libros as L where L.IdLibro=" & cLibros.fGetF("IdLibro") & " And L.IdLibro=r.IdLibro And r.IdAutor=Autors.IdAutor order by Autors.IdAutor")
'    cAutoresLocal.mPrimero
'    SiNo = MsgBox("¿Desea guardar la lectura en progreso?", vbQuestion + vbYesNo)
'    If SiNo = vbYes Then
'        If cAutoresLocal.fGetF("Autor_o_autores") = "" Then
'            cAutores.LetF txtAutor.Text, "Autor_o_autores"
'        End If
'    End If
End Sub

Private Sub cLibros_Actualizando()
    'On Error Resume Next
'    Set cAños = Nothing
'    Set cTemas = Nothing
'    Set cAutores = Nothing
'    Set cRelacion = Nothing
'
'
'    With cLibros
'        cmbNombreDelLibro.Text = .fGetF("NombreDelLibro")
'        txtPaginasLen.Text = .fGetF("CantidadDePaginas")
'        txtDireccionURL.Text = .fGetF("DireccionURL")
'        txtNotas.Text = .fGetF("Notas")
'    End With
    
'    'Se actualizan las tablas relacionadas.
'
'    Set cAños = New clsDbx
'    cAños.LoadDBConSQL ("Select a.AñoDePublicasion, a.IdAño from Años as a, Relacion r, Libros as L where L.IdLibro=" & cLibros.fGetF("IdLibro") & " And L.IdLibro=r.IdLibro And r.IdAño=a.IdAño order by a.IdAño")
'    cAños.mPrimero
'
'    Set cTemas = New clsDbx
'    cTemas.LoadDBConSQL ("Select t.Tema, t.IdTema from Temas as t, Relacion r, Libros as L where L.IdLibro=" & cLibros.fGetF("IdLibro") & " And L.IdLibro=r.IdLibro And r.IdTema=t.IdTema order by t.IdTema")
'    cTemas.mPrimero
'
'
'    Set cAutores = New clsDbx
'    cAutores.LoadDBConSQL ("Select IdRelacion from Relacion where IdLibro=" & cLibros.fGetF("IdLibro"))
'    cAutores.mPrimero
'
'
'    Set cRelacion = New clsDbx
'    cRelacion.LoadDBConSQL ("Select R.IdRelacion, R.IdTema, R.IdAño, R.IdLibro, R.IdAutor from Relacion as R Where R.IdLibro=" & cLibros.fGetF("IdLibro"))
'    cRelacion.mPrimero
End Sub



Private Sub cLibros_NuevoRegistro()
    cmbNombreDelLibro.Text = ""
    txtPaginasLen.Text = ""
    txtDireccionURL.Text = ""
    txtNotas.Text = ""
End Sub

Private Sub cLibros_SentFields()
'On Error Resume Next
'    Dim SiNo As Integer
'
'    SiNo = MsgBox("¿Desea guardar la lectura en progreso?", vbQuestion + vbYesNo)
'    If SiNo = vbYes Then
'        With cLibros
'            .LetF cmbNombreDelLibro.Text, "NombreDelLibro"
'            .LetF txtPaginasLen.Text, "CantidadDePaginas"
'            .LetF txtDireccionURL.Text, "DireccionURL"
'            .LetF txtNotas.Text, "Notas"
'        End With
'    End If
    
'    cRelacion = Nothing
'    Set cRelacion = New clsDbx
'    cRelacion.LoadDBConSQL ("Select R.IdRelacion, R.IdTema, R.IdAño, R.IdLibro, R.IdAutor from Relacion as R ") 'Where R.IdLibro=" & cLibros.fGetF("IdLibro"))
'    cRelacion.mPrimero
'    'De la siguiente manera me doy cuenta si existe la relacion.
'    lbId.Caption = cRelacion.fGetF("IdLibro")
End Sub

Private Sub cmbNombreDelLibro_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
    Select Case VarTipo(Source)
    Case "TextBox", "ComboBox"
        cmbNombreDelLibro.Text = Source '.Text
        Source.DragMode = 0
        Source.SelStart = 0
    Case "DbLectura"
        cmbNombreDelLibro.Text = Source.TxtTitulo
        Source.DragMode = 0
        Source.SeltxtTituloStart = 0
    End Select
End Sub

Private Sub CMDAnterior_Click()
    On Error GoTo N
'    If lbId.Caption = "" Then
'        DataEnvironment1.rsLibros.MoveFirst
'    End If
    If DataEnvironment1.rsLibros.BOF = False Then
        DataEnvironment1.rsLibros.MovePrevious
    Else
        DataEnvironment1.rsLibros.MoveFirst
    End If
    Exit Sub
N:
    On Error Resume Next
    If DataEnvironment1.rsLibros.BOF = False Then
        DataEnvironment1.rsLibros.MoveFirst
    End If
End Sub

Private Sub CMDAnterior_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro

End Sub

Private Sub CMDBuscar1_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub CMDEliminar_Click()
'En bases de datos no se deben borrar los datos.


End Sub

Private Sub CMDEliminar_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub CMDEnd_Click()
    Unload Me
End Sub

Private Sub CMDEnd_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub CMDGuardar_Click()
    DataEnvironment1.rsLibros.Update
    'DataEnvironment1.rsLibros.Requery
End Sub

Private Sub GuardarTablasRelacion() 'Tabla años.
'    Dim SiNo As Integer
'
'    Dim cAñosLocal As clsDbx
'    Set cAñosLocal = New clsDbx
'    cAñosLocal.LoadDBConSQL ("SELECT AñoDePublicasion FROM Años WHERE AñoDePublicasion='" & txtAñoDePublicasion.Text & "' order by AñoDePublicasion;") '("Select a.AñoDePublicasion from Años as a, Relacion r, Libros as L where L.IdLibro=" & cLibros.fGetF("IdLibro") & " And L.IdLibro=r.IdLibro And r.IdAño=a.IdAño order by a.IdAño")
'    cAñosLocal.mPrimero
'
'    Dim cAños1 As clsDbx
'    Set cAños1 = New clsDbx
'
'    cAños1.LoadDBs ("años") '("Slect AñoDePublicasion from Años where AñoDePublicasion=" & txtAñoDePublicasion.Text) '("Select a.AñoDePublicasion from Años as a, Relacion r, Libros as L where L.IdLibro=" & cLibros.fGetF("IdLibro") & " And L.IdLibro=r.IdLibro And r.IdAño=a.IdAño order by a.IdAño")
'
'    SiNo = True 'MsgBox("¿Desea guardar la lectura en progreso?", vbQuestion + vbYesNo)
'    If SiNo = vbYes Then
'    'MsgBox (cAñosLocal.fGetF("AñoDePublicasion"))
'        If cAñosLocal.fGetF("AñoDePublicasion") = "" Then
'            cAños1.Nuevo
'            cAños1.LetF txtAñoDePublicasion.Text, "AñoDePublicasion"
'            cAños1.mUltimo
'            MsgBox ("insertando")
'        Else
'            MsgBox ("No se pudo guardar la informacion " & cAñosLocal.fGetF("AñoDePublicasion"))
'        End If
'    End If
'
'    'Temas
'    Dim cTemasLocal As clsDbx
'    Set cTemasLocal = New clsDbx
'    cTemasLocal.LoadDBConSQL ("SELECT Tema FROM Temas WHERE Tema='" & txtTemaDelLibro.Text & "' order by Tema;") '("Select a.AñoDePublicasion from Años as a, Relacion r, Libros as L where L.IdLibro=" & cLibros.fGetF("IdLibro") & " And L.IdLibro=r.IdLibro And r.IdAño=a.IdAño order by a.IdAño")
'    cTemasLocal.mPrimero
'
'    Dim cTemas1 As clsDbx
'    Set cTemas1 = New clsDbx
'
'    cTemas1.LoadDBs ("Temas")   '("Slect AñoDePublicasion from Años where AñoDePublicasion=" & txtAñoDePublicasion.Text) '("Select a.AñoDePublicasion from Años as a, Relacion r, Libros as L where L.IdLibro=" & cLibros.fGetF("IdLibro") & " And L.IdLibro=r.IdLibro And r.IdAño=a.IdAño order by a.IdAño")
'
'    SiNo = True 'MsgBox("¿Desea guardar la lectura en progreso?", vbQuestion + vbYesNo)
'    If SiNo = vbYes Then
'    'MsgBox (cAñosLocal.fGetF("AñoDePublicasion"))
'        If cTemasLocal.fGetF("Tema") = "" Then
'            cTemas1.Nuevo
'            cTemas1.LetF txtTemaDelLibro.Text, "Tema"
'            cTemas1.mUltimo
'            MsgBox ("insertando")
'        Else
'            MsgBox ("No se pudo guardar la informacion " & cTemasLocal.fGetF("Tema"))
'        End If
'    End If
'
'    'Autores
'    Dim cAutoresLocal As clsDbx
'    Set cAutoresLocal = New clsDbx
'    cAutoresLocal.LoadDBConSQL ("SELECT Autor_o_autores FROM Autores WHERE Autor_o_autores='" & txtAutor.Text & "' order by Tema;") '("Select a.AñoDePublicasion from Años as a, Relacion r, Libros as L where L.IdLibro=" & cLibros.fGetF("IdLibro") & " And L.IdLibro=r.IdLibro And r.IdAño=a.IdAño order by a.IdAño")
'    cAutoresLocal.mPrimero
'
'    Dim cAutores1 As clsDbx
'    Set cAutores1 = New clsDbx
'
'    cAutores1.LoadDBs ("Autores") '("Slect AñoDePublicasion from Años where AñoDePublicasion=" & txtAñoDePublicasion.Text) '("Select a.AñoDePublicasion from Años as a, Relacion r, Libros as L where L.IdLibro=" & cLibros.fGetF("IdLibro") & " And L.IdLibro=r.IdLibro And r.IdAño=a.IdAño order by a.IdAño")
'
'    SiNo = True 'MsgBox("¿Desea guardar la lectura en progreso?", vbQuestion + vbYesNo)
'    If SiNo = vbYes Then
'    'MsgBox (cAñosLocal.fGetF("AñoDePublicasion"))
'        If cAutoresLocal.fGetF("Autor_o_autores") = "" Then
'            cAutores1.Nuevo
'            cAutores1.LetF txtAutor.Text, "Autor_o_autores"
'            cAutores1.mUltimo
'            MsgBox ("insertando")
'        Else
'            MsgBox ("No se pudo guardar la informacion " & cAutoresLocal.fGetF("Autor_o_autores"))
'        End If
'    End If
'
'    'Relacion
'    Dim cRelacionLocal As clsDbx
'    Set cRelacionLocal = New clsDbx
'    cRelacionLocal.LoadDBConSQL ("SELECT IdLibro FROM Relacion WHERE IdLibro='" & cLibros.fGetF("IdLibro") & "' order by IdLibro;") '("Select a.AñoDePublicasion from Años as a, Relacion r, Libros as L where L.IdLibro=" & cLibros.fGetF("IdLibro") & " And L.IdLibro=r.IdLibro And r.IdAño=a.IdAño order by a.IdAño")
'    cRelacionLocal.mPrimero
'
'    Dim cRelacion1 As clsDbx
'    Set cRelacion1 = New clsDbx
'
'    cRelacion1.LoadDBs ("Relacion") '("Slect AñoDePublicasion from Años where AñoDePublicasion=" & txtAñoDePublicasion.Text) '("Select a.AñoDePublicasion from Años as a, Relacion r, Libros as L where L.IdLibro=" & cLibros.fGetF("IdLibro") & " And L.IdLibro=r.IdLibro And r.IdAño=a.IdAño order by a.IdAño")
'
'    SiNo = True 'MsgBox("¿Desea guardar la lectura en progreso?", vbQuestion + vbYesNo)
'    If SiNo = vbYes Then
'    'MsgBox (cAñosLocal.fGetF("AñoDePublicasion"))
'        If cRelacionLocal.fGetF("IdLibro") = "" Then
'            cRelacion1.Nuevo
'            cRelacion1.LetF cLibros.fGetF("IdLibro"), "IdLibro"
'            cRelacion1.LetF cTemas1.fGetF("IdTema"), "IdTema"
'            cRelacion1.LetF cAños1.fGetF("IdAño"), "IdAño"
'            cRelacion1.LetF cAutores1.fGetF("IdAutor"), "IdAutor"
'            cRelacion1.mUltimo
'            MsgBox ("insertando")
'        Else
'            MsgBox ("No se pudo guardar la informacion " & cRelacionLocal.fGetF("IdLibro"))
'        End If
'    End Ifadas()
'
End Sub

Private Sub CMDGuardar_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub CMDNuevo_Click()
    
    cmbNombreDelLibro.Text = ""
    txtAutor.Text = ""
    txtPaginasLen.Text = ""
    txtAñoDePublicasion.Text = ""
    txtTemaDelLibro.Text = ""
    txtDireccionURL.Text = ""
    txtNotas.Text = ""
    DataEnvironment1.rsLibros.AddNew
End Sub

Private Sub CMDNuevo_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro

End Sub

Private Sub CMDPrimero_Click()
    On Error Resume Next
    If DataEnvironment1.rsLibros.BOF = False Then
        DataEnvironment1.rsLibros.MoveFirst
    End If
End Sub

Private Sub CMDPrimero_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub CMDSiguiente_Click()
    On Error GoTo N
'    If lbId.Caption = "" Then
'        DataEnvironment1.rsLibros.MoveLast
'    End If
    If DataEnvironment1.rsLibros.EOF = False Then
        DataEnvironment1.rsLibros.MoveNext
    Else
        DataEnvironment1.rsLibros.MoveLast
    End If
    Exit Sub
N:
    On Error Resume Next
    If DataEnvironment1.rsLibros.EOF = False Then
        DataEnvironment1.rsLibros.MoveLast
    End If
End Sub

Private Sub CMDSiguiente_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub CmdUltimo_Click()
    On Error Resume Next
    If DataEnvironment1.rsLibros.EOF = False Then
        DataEnvironment1.rsLibros.MoveLast
    End If
End Sub

Private Sub CmdUltimo_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
    
End Sub

Private Sub cRelacion_NuevoRegistro()
    lbId.Caption = ""
End Sub

Private Sub cRelacion_SentFields()
'    With cRelacion
'        .LetF cAutores.fGetF("IdAutor"), "IdAutor"
'        .LetF cAños.fGetF("IdAño"), "IdAño"
'        .LetF cLibros.fGetF("IdLibro"), "IdLibro"
'        .LetF cTemas.fGetF("IdTema"), "IdTema"
'    End With
End Sub


Private Sub cTemas_Actualizando()
'On Error Resume Next
'    txtTemaDelLibro.Text = cTemas.fGetF("Tema")
End Sub

Private Sub cTemas_NuevoRegistro()
    txtTemaDelLibro.Text = ""
End Sub

Private Sub cTemas_SentFields()
'On Error Resume Next
'    Dim SiNo As Integer
'    Dim cTemasLocal As clsDbx
'    Set cTemasLocal = New clsDbx
'    cTemasLocal.LoadDBConSQL ("Select tema from Temas where Tema=" & txtTemaDelLibro.Text) '("Select t.Tema from Temas as t, Relacion r, Libros as L where L.IdLibro=" & cLibros.fGetF("IdLibro") & " And L.IdLibro=r.IdLibro And r.IdTema=t.IdTema order by t.IdTema")
'    cTemasLocal.mPrimero
'    SiNo = MsgBox("¿Desea guardar la lectura en progreso?", vbQuestion + vbYesNo)
'    If SiNo = vbYes Then
'        If cTemasLocal.fGetF("Tema") = "" Then
'            cTemas.LetF txtTemaDelLibro.Text, "Tema"
'        End If
'    End If
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    Select Case VarTipo(Source)
    Case "TextBox", "ComboBox"
        cmbNombreDelLibro.Text = Source '.Text
        Source.DragMode = 0
        Source.SelStart = 0
    Case "DbLectura"
        cmbNombreDelLibro.Text = Source.TxtTitulo
        Source.DragMode = 0
        Source.SeltxtTituloStart = 0
    End Select

    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub Form_Load()
'    Set cLibros = New clsDbx
'    cLibros.LoadDBs "Libros"
'    cLibros.mPrimero
End Sub


Private Sub lbAñoDePublicasion_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub lbAutor_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub lbDireccionURL_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub lbId_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub lbNotas_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub lbPaginasLen_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub lbTemaDelLibro_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub txtAñoDePublicasion_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub txtAutor_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub txtDireccionURL_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub txtDireccionURL_GotFocus()
    txtDireccionURL.BackColor = Colores.Verde
End Sub

Private Sub txtDireccionURL_LostFocus()
    txtDireccionURL.BackColor = Colores.Defauld
End Sub

Private Sub txtNotas_Change()
    Dim L As Long
    L = Len(txtNotas)
    
    If L > pgbMemo.Max Then
        MsgBox "Solo puede escribir 62000 caracteres.", vbInformation
        Exit Sub
    End If
    
    pgbMemo.Value = L
    
    pgbMemo.ToolTipText = "Cantidad de caracteres: " & L & ". Maximo permitido 62000."

End Sub

Private Sub txtNotas_DragDrop(Source As Control, X As Single, Y As Single)
    'sAgarrarDatos Data, txtNotas
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub txtNotas_GotFocus()
    txtNotas.BackColor = Colores.Verde
End Sub

Private Sub txtNotas_LostFocus()
    txtNotas.BackColor = Colores.Defauld
End Sub

Private Sub txtNotas_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub txtPaginasLen_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub txtTemaDelLibro_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Public Sub sAgarrarDatos(Data As DataObject, ByRef Control_con_propiedad_text As Object) 'Obtiene el texto de otra ventana mediante el arrastre.
    On Error GoTo AccionesCorrectivas
    Control_con_propiedad_text.Text = Data.GetData(vbCFText)
    Exit Sub
AccionesCorrectivas:
    MsgBox "Tengo problemas con sAgarrarDatos"
End Sub
