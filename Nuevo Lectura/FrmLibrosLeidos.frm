VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLibrosLeidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libros leidos"
   ClientHeight    =   5310
   ClientLeft      =   1635
   ClientTop       =   1620
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   8580
   Begin VB.CommandButton CMDPrimero 
      Height          =   495
      Left            =   2610
      Picture         =   "FrmLibrosLeidos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Ir al primer registro"
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton CMDAnterior 
      Height          =   495
      Left            =   3090
      Picture         =   "FrmLibrosLeidos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Ir al registro anterior"
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton CMDSiguiente 
      Height          =   495
      Left            =   5010
      Picture         =   "FrmLibrosLeidos.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Ir al siguiente registro"
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton CmdUltimo 
      Height          =   495
      Left            =   5490
      Picture         =   "FrmLibrosLeidos.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Ir al ultimo registro (El más nuevo)"
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton CMDBuscar1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1560
      Picture         =   "FrmLibrosLeidos.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Buscar registros"
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton CMDNuevo 
      Height          =   495
      Left            =   120
      MouseIcon       =   "FrmLibrosLeidos.frx":0F32
      MousePointer    =   99  'Custom
      Picture         =   "FrmLibrosLeidos.frx":1084
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "0"
      ToolTipText     =   "Nuevo registro"
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton CMDGuardar 
      Height          =   495
      Left            =   840
      Picture         =   "FrmLibrosLeidos.frx":194E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Guardar registro"
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton CMDEliminar 
      Height          =   495
      Left            =   6345
      MouseIcon       =   "FrmLibrosLeidos.frx":2218
      MousePointer    =   99  'Custom
      Picture         =   "FrmLibrosLeidos.frx":2522
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Eliminar registro"
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton CMDEnd 
      Height          =   495
      Left            =   7920
      Picture         =   "FrmLibrosLeidos.frx":2DEC
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cerrar este formulario"
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox txtAutor 
      BackColor       =   &H00C0C0C0&
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
   Begin VB.ComboBox cmbNombreDelLibro 
      BackColor       =   &H00C0C0C0&
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
   Begin VB.Label lbId 
      BackColor       =   &H00808080&
      Caption         =   "lbId"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3585
      TabIndex        =   24
      Tag             =   "Id"
      Top             =   4800
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

Private Sub CMDAnterior_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub CMDBuscar1_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
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

Private Sub CMDGuardar_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub CMDNuevo_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub CMDPrimero_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub CMDSiguiente_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
End Sub

Private Sub CmdUltimo_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
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

Private Sub txtDireccionURL_Change()
    lbDireccionURL.BackColor = Colores.Defauld
End Sub

Private Sub txtDireccionURL_DragDrop(Source As Control, X As Single, Y As Single)
    'sArrastrarUnArchivo Data, txtDireccionURL, cmbNombreDelLibro
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
