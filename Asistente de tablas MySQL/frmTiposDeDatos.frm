VERSION 5.00
Begin VB.Form frmTiposDeDatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de datos"
   ClientHeight    =   2955
   ClientLeft      =   5355
   ClientTop       =   3540
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
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
      ItemData        =   "frmTiposDeDatos.frx":0000
      Left            =   600
      List            =   "frmTiposDeDatos.frx":0028
      TabIndex        =   2
      Top             =   1320
      Width           =   4815
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   2925
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   645
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seleccione el nuevo tipo de datos."
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lbTipoDeDatos 
      Caption         =   "Tipo de datos"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
End
Attribute VB_Name = "frmTiposDeDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private vTipoDeDatos As Variant
Private vNombreDelCampo As Variant




Private Sub cmdAceptar_Click()
    'Debug.Print prTipoDeDatos
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Public Property Get prTipoDeDatos() As Variant
    Dim frmChar As New frmValoresDelChar
    Select Case cmbTipoDeDatos.Text
    Case "double"
        vTipoDeDatos = " Default 0.0"
    Case "char"
        frmChar.Show vbModal, Me
        vTipoDeDatos = "char " & frmChar.Valor
    Case "year"
        vTipoDeDatos = "year default " & COMILLAS_DOBLES & " 1978" & COMILLAS_DOBLES
    Case "time"
        vTipoDeDatos = "time default " & COMILLAS_DOBLES & " 00:00:00" & COMILLAS_DOBLES
    Case "date"
        vTipoDeDatos = "date default " & COMILLAS_DOBLES & " 1978/1/25" & COMILLAS_DOBLES
    Case "DateTime"
        vTipoDeDatos = "DateTime default " & COMILLAS_DOBLES & " 1978/1/25 00:00:00" & COMILLAS_DOBLES
    Case "boolean"
        vTipoDeDatos = " default false"
    Case "Primary Key"
        vTipoDeDatos = " int not null auto_increment, primary key (" & vNombreDelCampo & ")"
    Case "Blob(imagen o archivo)"
        vTipoDeDatos = " blob  NOT NULL"
    Case "LongBlob(imagenes o archivos grandes)"
        vTipoDeDatos = " LongBlob  NOT NULL"
    Case "foreign key"
        vTipoDeDatos = " int , foreign key (" & vNombreDelCampo & ")"
    Case Else
        vTipoDeDatos = cmbTipoDeDatos.Text
    End Select
    prTipoDeDatos = vTipoDeDatos
End Property

Public Property Let prTipoDeDatos(vNewValue As Variant)
    cmbTipoDeDatos.Text = vNewValue
End Property

Public Property Get prNombreDelCampo() As Variant
    prNombreDelCampo = vNombreDelCampo
End Property

Public Property Let prNombreDelCampo(vNewValue As Variant)
    vNombreDelCampo = vNewValue
End Property

