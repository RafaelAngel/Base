VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCreadorDeTablas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asistente para crear tablas MySQL"
   ClientHeight    =   7365
   ClientLeft      =   435
   ClientTop       =   930
   ClientWidth     =   11340
   Icon            =   "frmCreadorDeTablas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11340
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   5640
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtComentario 
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   4440
      Width           =   10935
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Borrar"
      Height          =   615
      Left            =   9720
      TabIndex        =   14
      ToolTipText     =   "Borra el codigo MySQL de creacion de tablas."
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CheckBox chkAvisar 
      Caption         =   "Lanzar msgBox despues de crear la tabla."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   13
      Top             =   0
      Value           =   1  'Checked
      Width           =   3975
   End
   Begin VB.CommandButton cmdCrear 
      Caption         =   "Crear"
      Enabled         =   0   'False
      Height          =   615
      Left            =   3000
      TabIndex        =   12
      ToolTipText     =   "Crear tabla"
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   615
      Left            =   4680
      TabIndex        =   10
      Top             =   8400
      Width           =   2055
   End
   Begin VB.TextBox txtTablaCreada 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   6000
      Width           =   10935
   End
   Begin VB.ListBox lstTiposDeDatos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   3960
      TabIndex        =   8
      Top             =   2160
      Width           =   7215
   End
   Begin VB.ComboBox cmbTipoDeDatos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmCreadorDeTablas.frx":08CA
      Left            =   3960
      List            =   "frmCreadorDeTablas.frx":08F2
      TabIndex        =   6
      Top             =   1560
      Width           =   4815
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9000
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtNombreDelCampo 
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
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   3495
   End
   Begin VB.ListBox lstNombresDeLosCampos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   3495
   End
   Begin VB.TextBox txtNombreDeLaTabla 
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lbCantidadDeCampos 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cantidad de campos=0"
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      Tag             =   "Cantidad de campos=0"
      Top             =   5160
      Width           =   3615
   End
   Begin VB.Label lbComentario 
      Caption         =   "Comentario descriptivo de la tabla, sirve de ayuda para saver la funcion de la tabla que se esta creando."
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4200
      Width           =   10935
   End
   Begin VB.Label lbTablaCreada 
      Caption         =   "Tabla por crear"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label lbTipoDeDatos 
      Caption         =   "Tipo de datos"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lbNombreDelCampo 
      Caption         =   "Nombre del campo"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label lbNombreDeLaBase 
      Caption         =   "Nombre de la tabla"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuConsultas 
      Caption         =   "Consultas"
      Begin VB.Menu mnuConsultaDeDatos 
         Caption         =   "Consulta de datos"
      End
      Begin VB.Menu mnu_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConsultaInsert 
         Caption         =   "Consulta insert"
      End
      Begin VB.Menu mnu_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConsultaUpdate 
         Caption         =   "Consulta Update"
      End
      Begin VB.Menu mnu_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConsultaDelete 
         Caption         =   "Consulta Delete"
      End
   End
   Begin VB.Menu mnuHerramientas 
      Caption         =   "Herramientas"
      Begin VB.Menu mnuProcedimientoAlmacenado 
         Caption         =   "Crear procedimiento almacemado"
      End
      Begin VB.Menu mnu_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFuncionAlmacenada 
         Caption         =   "Crear funcion almacenada"
      End
      Begin VB.Menu mnu_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCrearDisparador 
         Caption         =   "Crear disparador"
      End
      Begin VB.Menu mnu_6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCrearBaseDeDatos 
         Caption         =   "Crear base de datos"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mnuAcercaDeRafaelMF 
         Caption         =   "Acerca de Rafael Angel montero Fernández"
      End
   End
End
Attribute VB_Name = "frmCreadorDeTablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Autor Rafael Angel Montero Fernández
'Correo: Sharkyc12@gmail.com
'El formato blop es para almacenar imagenes o archivos.
Private vfrmTipoDeDatos As New frmTiposDeDatos
Private Type tyPasosDeEvolucion
    vNombre_de_la_tabla_escrito As Integer
    vNombre_del_campo_escrito As Integer
    vTipo_de_datos_seleccionado As Integer
    vSe_agrego_el_nuevo_campo_a_la_lista As Integer
End Type
Private vPasosDeLaCreacionDeTabla As tyPasosDeEvolucion

Private frmChar As New frmValoresDelChar
Attribute frmChar.VB_VarHelpID = -1

Private Type tyTabla
    vNombre_de_la_tabla As Variant
    mNombres_de_los_campos() As String
    mTipos_de_datos() As String
    vNombres_de_los_campos As Variant
    vTipos_de_datos As Variant 'Incluye el tamaño y valor por defecto.
    vValor As Variant 'Sirve para guardar momentaneamente el valor del campo.
End Type
Private vTabla As tyTabla

Private Sub activarCrearTabla()
    If txtNombreDeLaTabla <> "" Then
        vPasosDeLaCreacionDeTabla.vNombre_de_la_tabla_escrito = 1
    Else
        vPasosDeLaCreacionDeTabla.vNombre_de_la_tabla_escrito = 0
    End If

    If txtNombreDeLaTabla <> "" And Me.lstNombresDeLosCampos.ListCount > 0 And Me.lstTiposDeDatos.ListCount > 0 Then
        cmdCrear.Enabled = True
    Else
        cmdCrear.Enabled = False
        'vPasosDeLaCreacionDeTabla.vNombre_de_la_tabla_escrito = 0
    End If
    evolucion
End Sub

Private Sub cmbTipoDeDatos_Click()
    cmbTipoDeDatos.Tag = cmbTipoDeDatos.Text
    Select Case cmbTipoDeDatos.Text
    Case "double"
        vTabla.vValor = " Default 0.0"
    Case "char"
        frmChar.Show vbModal, Me
        vTabla.vValor = frmChar.Valor
    Case "year"
        vTabla.vValor = " default " & COMILLAS_DOBLES & " 1978" & COMILLAS_DOBLES
    Case "time"
        vTabla.vValor = " default " & COMILLAS_DOBLES & " 00:00:00" & COMILLAS_DOBLES
    Case "date"
        vTabla.vValor = " default " & COMILLAS_DOBLES & " 1978/1/25" & COMILLAS_DOBLES
    Case "DateTime"
        vTabla.vValor = " default " & COMILLAS_DOBLES & " 1978/1/25 00:00:00" & COMILLAS_DOBLES
    Case "boolean"
        vTabla.vValor = " default false"
    Case "Primary Key"
        vTabla.vValor = " int not null auto_increment, primary key (" & Me.txtNombreDelCampo.Text & ")"
    Case "Blob(imagen o archivo)"
        vTabla.vValor = " blob NULL Default NULL"
    Case "LongBlob(imagenes o archivos grandes)"
        vTabla.vValor = " LongBlob NULL Default NULL"
    Case "foreign key"
        Dim vNombre_de_la_tabla_foranea As Variant
        vNombre_de_la_tabla_foranea = InputBox("Escriba el nombre de la tabla.", "Creacion de clave foranea", "Table")
        If vNombre_de_la_tabla_foranea = "" Then
            vNombre_de_la_tabla_foranea = "Table"
        End If
        vTabla.vValor = " int, foreign key (" & Me.txtNombreDelCampo.Text & ")References " & vNombre_de_la_tabla_foranea
    End Select
    activarAgregarCampo
End Sub



Private Sub activarAgregarCampo()
    If Me.txtNombreDelCampo.Text <> "" Then
        cmdAgregar.Enabled = True
        vPasosDeLaCreacionDeTabla.vNombre_del_campo_escrito = 1
        cmbTipoDeDatos.Enabled = True
        
    Else
        cmbTipoDeDatos.Enabled = False
    
        cmdAgregar.Enabled = False
        'vPasosDeLaCreacionDeTabla.vNombre_del_campo_escrito = 0
    End If
    If cmbTipoDeDatos.Text <> "" Then
        cmdAgregar.Enabled = True
        vPasosDeLaCreacionDeTabla.vTipo_de_datos_seleccionado = 1
    Else
        cmdAgregar.Enabled = False
        'vPasosDeLaCreacionDeTabla.vTipo_de_datos_seleccionado = 0
    End If
    evolucion
End Sub

Private Sub evolucion()
'On Error Resume Next
        'ProgressBarTabla.Value = vPasosDeLaCreacionDeTabla.vNombre_de_la_tabla_escrito + vPasosDeLaCreacionDeTabla.vNombre_del_campo_escrito + vPasosDeLaCreacionDeTabla.vTipo_de_datos_seleccionado + vPasosDeLaCreacionDeTabla.vSe_agrego_el_nuevo_campo_a_la_lista
End Sub

Private Sub cmdAgregar_Click()
    ProgressBar1.Value = 0
    lstNombresDeLosCampos.AddItem Me.txtNombreDelCampo
    txtNombreDelCampo.Text = ""
    Select Case cmbTipoDeDatos.Text
    Case "Primary Key"
        lstTiposDeDatos.AddItem vTabla.vValor
    Case "Blob(imagen o archivo)"
        lstTiposDeDatos.AddItem vTabla.vValor
    Case "LongBlob(imagenes o archivos grandes)"
        lstTiposDeDatos.AddItem vTabla.vValor
    Case "foreign key"
        lstTiposDeDatos.AddItem vTabla.vValor
    Case Else
    'MsgBox "Extra"
        Me.lstTiposDeDatos.AddItem Me.cmbTipoDeDatos.Text & vTabla.vValor
    End Select
    
    cmbTipoDeDatos.Text = ""
    vTabla.vValor = ""
    activarCrearTabla
    vPasosDeLaCreacionDeTabla.vSe_agrego_el_nuevo_campo_a_la_lista = 1
    evolucion
    lbTablaCreada.Caption = "Tabla por crear"
    lbCantidadDeCampos.Caption = "Cantidad de campos= " & Me.lstNombresDeLosCampos.ListCount
End Sub

Private Sub cmdAgregar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    activarAgregarCampo
End Sub

Private Sub cmdClear_Click()
    ProgressBar1.Value = 0
    lbCantidadDeCampos.Caption = lbCantidadDeCampos.Tag
    txtTablaCreada.Text = ""
    txtComentario.Text = ""
    With vPasosDeLaCreacionDeTabla
        .vNombre_de_la_tabla_escrito = 0
        .vNombre_del_campo_escrito = 0
        .vTipo_de_datos_seleccionado = 0
        .vSe_agrego_el_nuevo_campo_a_la_lista = 0
    End With
    evolucion
    lbTablaCreada.Caption = "Tabla por crear"
End Sub

Private Sub CmdCrear_Click()
    Dim i As Integer
    Dim vCampos As Variant
    On Error GoTo S
    ProgressBar1.Max = lstNombresDeLosCampos.ListCount - 1
S:
    For i = 0 To Me.lstNombresDeLosCampos.ListCount - 1
        ProgressBar1.Value = i
        lstNombresDeLosCampos.ListIndex = i
        Me.lstTiposDeDatos.ListIndex = i
        If i < lstNombresDeLosCampos.ListCount - 1 Then
            vCampos = vCampos & lstNombresDeLosCampos.Text & " " & lstTiposDeDatos.Text & ", "
        ElseIf i = lstNombresDeLosCampos.ListCount - 1 Then
            vCampos = vCampos & lstNombresDeLosCampos.Text & " " & lstTiposDeDatos.Text
        End If
    Next
    Dim vComentario As Variant
    If txtComentario.Text <> "" Then
        vComentario = "#" & txtComentario.Text & rtc
    End If
    Me.txtTablaCreada.Text = vComentario & "Drop table if exists " & Me.txtNombreDeLaTabla.Text & ";" & rtc & "create table " & txtNombreDeLaTabla.Text & "(" & vCampos & ");"
    Clipboard.Clear
    Clipboard.SetText Me.txtTablaCreada.Text
    cmdCrear.Enabled = False
    lstNombresDeLosCampos.Clear
    lstTiposDeDatos.Clear
    txtNombreDeLaTabla.Text = ""
    lbTablaCreada.Caption = "Tabla creada"
    If chkAvisar.Value = 1 Then
        MsgBox "Datos pegados en el portapapeles", vbInformation
    End If
End Sub

Private Sub cmdSalir_Click()
    On Error Resume Next
    Unload Me
End Sub






Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    activarAgregarCampo
    activarCrearTabla
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub lbNombreDeLaBase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
activarAgregarCampo
    activarCrearTabla
End Sub

Private Sub lbNombreDelCampo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    activarAgregarCampo
    activarCrearTabla
End Sub



Private Sub lbTipoDeDatos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    activarAgregarCampo
    activarCrearTabla
End Sub

Private Sub lstNombresDeLosCampos_Click()
    Me.lstTiposDeDatos.ListIndex = lstNombresDeLosCampos.ListIndex
End Sub

Private Sub lstNombresDeLosCampos_DblClick()
    Dim vNuevaDefinicion As Variant
    vNuevaDefinicion = InputBox("Escriba la nueva definicion", , lstNombresDeLosCampos.Text)
    If vNuevaDefinicion <> "" Then
        lstNombresDeLosCampos.List(lstNombresDeLosCampos.ListIndex) = vNuevaDefinicion
    End If
End Sub

Private Sub lstTiposDeDatos_Click()
    cmbTipoDeDatos.Enabled = False
    lstNombresDeLosCampos.ListIndex = lstTiposDeDatos.ListIndex
End Sub


Private Sub lstTiposDeDatos_DblClick()
    Dim vNuevaDefinicion As Variant
    vfrmTipoDeDatos.prTipoDeDatos = cmbTipoDeDatos.Tag
    vfrmTipoDeDatos.prNombreDelCampo = lstNombresDeLosCampos.Text
    vfrmTipoDeDatos.Show vbModal, Me
    vNuevaDefinicion = vfrmTipoDeDatos.prTipoDeDatos
    If vNuevaDefinicion <> "" Then
        lstTiposDeDatos.List(lstTiposDeDatos.ListIndex) = vNuevaDefinicion
    End If
    Unload vfrmTipoDeDatos
End Sub

Private Sub lstTiposDeDatos_LostFocus()
    cmbTipoDeDatos.Enabled = True
End Sub

Private Sub mnuAcercaDeRafaelMF_Click()
    MsgBox "Autor de este programa Asistente para crear tablas en lenguaje MySQL" & rtc & "Rafael Angel Montero Fernández" & rtc & "Correo: Sharkyc12@Gmail.com" & rtc & "Celular: 83942235"
End Sub

Private Sub mnuConsultaDeDatos_Click()
    frmConsultaSelect.Show
End Sub

Private Sub mnuConsultaInsert_Click()
    frmInsert.Show
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub

Private Sub txtNombreDeLaTabla_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
activarAgregarCampo
    activarCrearTabla
End Sub

Private Sub txtNombreDelCampo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    activarAgregarCampo
    activarCrearTabla
    
End Sub
