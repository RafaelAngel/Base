VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCrearScript 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Esperando nombre de la tabla"
   ClientHeight    =   7425
   ClientLeft      =   780
   ClientTop       =   885
   ClientWidth     =   10350
   Icon            =   "frmPruebaDConexion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   10350
   Begin VB.CommandButton CmdCopy 
      Caption         =   "Copiar"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Copia la script al portapapeles"
      Top             =   2520
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar PgbFields 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtDatos 
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin RichTextLib.RichTextBox rtxtDatosProcesados 
      Height          =   975
      Left            =   3720
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmPruebaDConexion.frx":324A
   End
   Begin RichTextLib.RichTextBox rtxtInsert 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   6165
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmPruebaDConexion.frx":32CC
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2880
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3480
      Top             =   2280
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   5520
      TabIndex        =   0
      Top             =   1080
      Width           =   4815
   End
   Begin MSComctlLib.ProgressBar PgbDatos 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label LbListaDeCampos 
      Caption         =   "Lista de campos de la tabla"
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label LbNombreDeLaTabla 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre de la tabla"
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
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmCrearScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private db As New clsDatos


Public activar As Boolean


Public Sub clear()
    List1.clear
End Sub

Public Sub add(Nombre As Variant)
    List1.AddItem Nombre
End Sub

Public Sub CargarDatos()
On Error Resume Next
    Me.SetFocus
    
    Dim insert As Variant
    db.Cerrar
    db.LoadDBs Me.Caption, direccion
    LbNombreDeLaTabla.Caption = Me.Caption
    db.mPrimero
    Me.txtDatos.Text = db.GetF(1)
    'cmdIrAdelante.Enabled = True
    Dim i As Integer
    Dim vFields As Variant
    
    List1.Text = "Id"
    If LCase(List1.Text) = LCase("id") Then
        List1.RemoveItem (List1.ListIndex)
    End If
    'For i = 0 To List1.ListCount - 1
        Dim r As Integer
        Dim vDatosCount As Integer
        vDatosCount = db.Cont - 1
        PgbFields.Max = vDatosCount
        For r = 0 To vDatosCount
            Dim c As Integer
            Dim vListaDeCamposCount As Integer
            vListaDeCamposCount = List1.ListCount - 1
            PgbDatos.Max = vListaDeCamposCount
            For c = 0 To vListaDeCamposCount
                List1.ListIndex = c
                
                If c < List1.ListCount - 1 Then
                    rtxtDatosProcesados.Text = rtxtDatosProcesados.Text & COMILLAS_DOBLES & db.GetF(List1.Text) & COMILLAS_DOBLES & ","
                    vFields = vFields & List1.Text & ","
                ElseIf c >= List1.ListCount - 1 Then
                    rtxtDatosProcesados.Text = rtxtDatosProcesados.Text & COMILLAS_DOBLES & db.GetF(List1.Text) & COMILLAS_DOBLES
                    vFields = vFields & List1.Text
                End If
                PgbDatos.Value = c
            Next
            db.mSiguiente
            rtxtInsert.Text = rtxtInsert.Text & "insert into " & Me.Caption & "(" & vFields & ")values(" & rtxtDatosProcesados.Text & ");" & rtc & rtc
            vFields = ""
            rtxtDatosProcesados.Text = ""
            PgbFields.Value = r
        Next
    'Next
    Dim Escribir As New ClsUsoDeArchivos
    Escribir.prDireccionParaElArchivo = vDireccionParaElArchivoDeTexto
    Escribir.vNombreDelArchivo = vNombreDelScript & " insercion de datos.sql"
    Escribir.EscribirDatos (rtxtInsert.Text)
    Me.WindowState = 1 'Minimizado
End Sub



Private Sub CmdCopy_Click()
Clipboard.clear
Clipboard.SetText (rtxtInsert.Text)
End Sub

Private Sub Timer1_Timer()
    CargarDatos
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
DoEvents
    If activar = True Then
        Timer1.Enabled = True
        Timer2.Enabled = False
    End If
End Sub


