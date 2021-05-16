VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPrincipal 
   Caption         =   "Traductor db simple"
   ClientHeight    =   7080
   ClientLeft      =   1050
   ClientTop       =   1200
   ClientWidth     =   10740
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   10740
   Begin VB.CommandButton CmdEscribir 
      Caption         =   "Escribir Script"
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   4895
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmPrincipal.frx":324A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBarTablas 
      Height          =   255
      Left            =   7920
      TabIndex        =   8
      Top             =   2160
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBarFields 
      Height          =   255
      Left            =   7920
      TabIndex        =   7
      Top             =   1680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   735
      Left            =   4643
      TabIndex        =   6
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox txtNombreDeLaBase 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Width           =   5655
   End
   Begin VB.TextBox txtDireccion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   8415
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "Abrir base"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Base de datos Access (*.mdb)|*.mdb"
   End
   Begin VB.Label Label1 
      Caption         =   "Script MySQL para la creacion de tablas"
      Height          =   375
      Left            =   3323
      TabIndex        =   5
      Top             =   2760
      Width           =   4095
   End
   Begin VB.Label lbNombre 
      Caption         =   "Nombre de la base"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Label lbDireccion 
      Caption         =   "Direccion con nombre de la base"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents clsMySQL As clsCrearMySQL
Attribute clsMySQL.VB_VarHelpID = -1
Private vDatos As Variant
Private WithEvents Escribir As ClsUsoDeArchivos
Attribute Escribir.VB_VarHelpID = -1

Private Sub clsMySQL_eveCrearInsertDatos(linea_sql_insert As Variant)
'On Error Resume Next
'Text3.SelStart = Len(Text2.Text)
vDatos = vDatos & rtc & linea_sql_insert & rtc & rtc
End Sub

Private Sub clsMySQL_eveCrearTablaMySQL(linea_sql_de_la_tabla As Variant)
RichTextBox1.Text = RichTextBox1.Text & linea_sql_de_la_tabla
End Sub

Private Sub clsMySQL_eveTrabajandoEnCampos(conteo As Integer, maximo_de_tareas As Integer)
    On Error Resume Next
    Me.ProgressBarFields.Max = maximo_de_tareas
    ProgressBarFields.Value = conteo
End Sub

Private Sub clsMySQL_eveTrabajandoEnTablas(conteo As Integer, maximo_de_tareas As Integer)
    On Error Resume Next
    Me.ProgressBarTablas.Max = maximo_de_tareas
    ProgressBarTablas.Value = conteo
End Sub

Private Sub cmdAbrir_Click()
'On Error GoTo n
Me.CommonDialog1.ShowOpen
Me.Caption = "Traduciendo las tablas de " & CommonDialog1.FileTitle
CrearCarpeta (Left(CommonDialog1.FileName, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle)))
Me.txtNombreDeLaBase.Text = CommonDialog1.FileTitle
Me.txtDireccion.Text = CommonDialog1.FileName 'Left(CommonDialog1.FileName, Len(CommonDialog1.FileName) - Len(txtNombreDeLaBase.Text))
Dim vNombreDeLaBase As String
vNombreDeLaBase = Left(txtNombreDeLaBase.Text, (Len(txtNombreDeLaBase.Text) - 4))
vNombreDelScript = vNombreDeLaBase
RichTextBox1.Text = "Drop database if exists " & vNombreDeLaBase & ";" & rtc
RichTextBox1.Text = RichTextBox1.Text & "Create database " & vNombreDeLaBase & ";" & rtc
RichTextBox1.Text = RichTextBox1.Text & "Use " & vNombreDeLaBase & ";" & rtc
RichTextBox1.Text = RichTextBox1.Text & "Set Autocommit=0;" & rtc
RichTextBox1.Text = RichTextBox1.Text & "Set MySQl_Safe_Updates=0;" & rtc & rtc

clsMySQL.crearMySQLBase CommonDialog1.FileName
Clipboard.clear

RichTextBox1.Text = RichTextBox1.Text & rtc & rtc & vDatos
'Clipboard.SetText RichTextBox1.Text
CmdEscribir.Enabled = True
'MsgBox "Se ha finalizado la traduccion.", vbInformation
'Exit Sub
'n:
    'RichTextBox1.Text = "Accion de traduccion de la base de datos cancelada."
End Sub




Private Sub CmdEscribir_Click()
'Dim Escribir As New ClsUsoDeArchivos
Escribir.prDireccionParaElArchivo = vDireccionParaElArchivoDeTexto
Dim Leer As Variant
Leer = ""
Escribir.vNombreDelArchivo = Left(txtNombreDeLaBase.Text, Len(txtNombreDeLaBase.Text) - 4) & " Base.sql"
Escribir.EscribirDatos (RichTextBox1.Text)
'Luego arreglo esto.
'Escribir.LeerTexto ("CodigoMysql.sql")

Dim vDatos As Variant
vDatos = Me.RichTextBox1.Text & Leer

End Sub


Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub CrearCarpeta(Direccion_para_la_carpeta As String)
    On Error Resume Next
    'MsgBox (Direccion_para_la_carpeta)
    vDireccionParaElArchivoDeTexto = Direccion_para_la_carpeta
EliminandoCarpeta
    MkDir Direccion_para_la_carpeta & "Script\" 'Crea una nueva carpeta.
End Sub


Private Sub EliminandoCarpeta()
    On Error Resume Next
    Kill vDireccionParaElArchivoDeTexto & "Script\*.*" 'Borra cuanto archivo exista.
'EliminarCarpeta:
    'On Error GoTo CrearCarpeta
    RmDir vDireccionParaElArchivoDeTexto & "Script\" 'Borra la carpeta por si existe y es una segiunda pasada.
'CrearCarpeta:
End Sub


Private Sub Escribir_EveLinea(Datos As Variant)
Escribir.EscribirDatos (Datos)
End Sub

Private Sub Form_Load()
    Set clsMySQL = New clsCrearMySQL
    Set Escribir = New ClsUsoDeArchivos
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
